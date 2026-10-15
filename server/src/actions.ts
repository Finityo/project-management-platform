import { defineAction, z, type ActionsModule } from "@hatch/space-sdk";
import { and, asc, desc, eq } from "drizzle-orm";
import * as schema from "./schema";
import { createBlankWorkbookBytes, createProjectWorkbookBytes, safeFileStem } from "./workbook";

const projectShape = z.object({
  id: z.number(), projectCode: z.string(), name: z.string(), sponsor: z.string(),
  projectManager: z.string(), businessOwner: z.string(), projectDate: z.string(),
  version: z.string(), overview: z.string(), updatedAt: z.string(),
});
const mutationResponse = z.object({ ok: z.boolean(), message: z.string() });
const workbookProjectShape = z.object({
  projectCode: z.string().trim().min(1).max(40), name: z.string().trim().min(1).max(120),
  sponsor: z.string().max(120), projectManager: z.string().max(120), businessOwner: z.string().max(120),
  projectDate: z.string().max(40), version: z.string().max(30), overview: z.string().max(4000),
});
const workbookSectionShape = z.object({ moduleKey: z.string().min(1).max(60), dataJson: z.string().max(50000) });
const workbookRowShape = z.object({ moduleKey: z.string().min(1).max(60), dataJson: z.string().max(50000), sortOrder: z.number().int().min(0) });
const importWorkbookResponse = z.object({
  ok: z.boolean(), id: z.number().optional(), projectCode: z.string().optional(), message: z.string(),
  sectionsCreated: z.number(), rowsCreated: z.number(),
});
const prepareWorkbookResponse = z.object({
  ok: z.boolean(), url: z.string().optional(), fileName: z.string().optional(), message: z.string(),
});
const gapShape = z.object({
  type: z.enum(["requirement", "wbs", "task"]),
  label: z.string(),
  detail: z.string(),
  moduleKey: z.string(),
  rowId: z.number(),
});
const traceabilityResponse = z.object({
  ok: z.literal(true),
  message: z.string(),
  checkedAt: z.string(),
  requirementsChecked: z.number(),
  wbsChecked: z.number(),
  tasksChecked: z.number(),
  gaps: z.array(gapShape),
});
const reportResponse = z.object({ ok: z.literal(true), message: z.string(), draft: z.string() });
const raciIssueShape = z.object({ rowId: z.number(), activity: z.string(), issue: z.enum(["no one accountable", "conflicting accountability"]), accountableCount: z.number() });
const raciResponse = z.object({
  ok: z.literal(true), message: z.string(), checkedAt: z.string(), activitiesChecked: z.number(), issues: z.array(raciIssueShape),
});
const scheduleRiskShape = z.object({
  rowId: z.number(), moduleKey: z.enum(["task-tracker", "timeline"]), taskName: z.string(), reason: z.string(), relevantDate: z.string(),
});
const scheduleHealthResponse = z.object({
  ok: z.literal(true), message: z.string(), checkedAt: z.string(), tasksChecked: z.number(), timelineItemsChecked: z.number(), risks: z.array(scheduleRiskShape),
});
const portfolioBudgetItemShape = z.object({
  projectId: z.number(), projectName: z.string(), projectCode: z.string(), budgeted: z.number(), actual: z.number(), variance: z.number(), hasBudgetData: z.boolean(),
});
const portfolioHealthShape = z.object({
  projectId: z.number(), projectName: z.string(), projectCode: z.string(), health: z.enum(["Red", "Amber", "Green"]), reasons: z.array(z.string()),
  percentComplete: z.number(), taskCount: z.number(), targetDate: z.string(), daysToTarget: z.number().nullable(),
  budgeted: z.number(), actual: z.number(), variance: z.number(),
  attention: z.object({ traceability: z.number(), raci: z.number(), schedule: z.number() }),
});
const portfolioDateShape = z.object({
  projectId: z.number(), projectName: z.string(), projectCode: z.string(), title: z.string(), date: z.string(), type: z.enum(["Task", "Milestone", "Key date", "Work package"]), owner: z.string(), overdue: z.boolean(),
});
const portfolioRiskShape = z.object({
  projectId: z.number(), projectName: z.string(), projectCode: z.string(), kind: z.enum(["Blocker", "Assumption", "Constraint"]), title: z.string(), detail: z.string(), owner: z.string(), impact: z.string(),
});
const portfolioResponse = z.object({
  ok: z.literal(true), refreshedAt: z.string(), asOfDate: z.string(),
  totals: z.object({
    projects: z.number(), red: z.number(), amber: z.number(), green: z.number(),
    budgeted: z.number(), actual: z.number(), variance: z.number(),
    taskStatus: z.object({ notStarted: z.number(), inProgress: z.number(), complete: z.number(), blocked: z.number() }),
  }),
  projects: z.array(portfolioHealthShape), budgets: z.array(portfolioBudgetItemShape), dates: z.array(portfolioDateShape), risks: z.array(portfolioRiskShape),
});

type RowData = Record<string, string>;

function parseRowData(value: string): RowData {
  try { return JSON.parse(value) as RowData; } catch { return {}; }
}

function splitReferences(value: string | undefined): string[] {
  return (value ?? "").split(/[,;\n]/).map((item) => item.trim().toLowerCase()).filter(Boolean);
}

function displayTask(data: RowData, fallback: string): string {
  return data.task?.trim() || data.taskName?.trim() || data.deliverable?.trim() || data.description?.trim() || fallback;
}

function rowText(data: RowData, ...keys: string[]): string {
  for (const key of keys) {
    const value = data[key];
    if (value !== undefined && value !== null && String(value).trim()) return String(value).trim();
  }
  return "";
}

function rowNumber(data: RowData, ...keys: string[]): number {
  const value = rowText(data, ...keys).replace(/[$,]/g, "");
  const parsed = Number(value || 0);
  return Number.isFinite(parsed) ? parsed : 0;
}

function isCompleteRow(data: RowData): boolean {
  const status = rowText(data, "status").toLowerCase();
  return ["complete", "completed", "done", "closed", "resolved"].includes(status) || rowNumber(data, "percentComplete", "percent_complete") >= 100;
}

function hasMeaningfulBlocker(data: RowData): boolean {
  const blocker = rowText(data, "blocker", "risk", "blockerRisk");
  return Boolean(blocker) && !["none", "n/a", "na", "no"].includes(blocker.toLowerCase());
}

const workbookSectionKeys = new Set(["business-case", "charter", "scope", "management-plan", "raci"]);
const workbookRowKeys = new Set(["charter", "stakeholders", "objectives", "requirements", "ac-log", "kickoff", "raci", "wbs", "pert", "timeline", "task-tracker", "todo", "calendar"]);

function validatedWorkbookJson(value: string): string | null {
  try {
    const parsed = JSON.parse(value) as unknown;
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return null;
    const clean: RowData = {};
    for (const [key, cell] of Object.entries(parsed)) clean[key] = cell === null || cell === undefined ? "" : String(cell);
    return JSON.stringify(clean);
  } catch {
    return null;
  }
}

export const Actions = {
  listProjects: defineAction({
    request: z.object({}),
    response: z.object({ projects: z.array(projectShape) }),
    async handler(ctx) {
      const db = ctx.db<typeof schema>();
      const rows = await db.select().from(schema.projects).orderBy(desc(schema.projects.updatedAt));
      return { projects: rows.map((p) => ({ ...p, updatedAt: p.updatedAt.toISOString() })) };
    },
  }),

  prepareWorkbookExport: defineAction({
    request: z.object({ projectId: z.number().int().positive().optional(), template: z.boolean() }),
    response: prepareWorkbookResponse,
    async handler(ctx, args): Promise<z.infer<typeof prepareWorkbookResponse>> {
      const db = ctx.db<typeof schema>();
      let bytes: Uint8Array;
      let fileName: string;
      if (args.template) {
        bytes = createBlankWorkbookBytes();
        fileName = "project-workbook-template.xlsx";
      } else {
        if (!args.projectId) return { ok: false, message: "Choose a project before preparing an export." };
        const [projectRows, sections, rows] = await Promise.all([
          db.select().from(schema.projects).where(eq(schema.projects.id, args.projectId)).limit(1),
          db.select().from(schema.sections).where(eq(schema.sections.projectId, args.projectId)),
          db.select().from(schema.moduleRows).where(eq(schema.moduleRows.projectId, args.projectId)).orderBy(asc(schema.moduleRows.sortOrder), asc(schema.moduleRows.id)),
        ]);
        const project = projectRows[0];
        if (!project) return { ok: false, message: "The selected project could not be found." };
        bytes = createProjectWorkbookBytes({
          project: {
            projectCode: project.projectCode, name: project.name, sponsor: project.sponsor,
            projectManager: project.projectManager, businessOwner: project.businessOwner,
            projectDate: project.projectDate, version: project.version, overview: project.overview,
          },
          sections: sections.map((section) => ({ moduleKey: section.moduleKey, dataJson: section.dataJson })),
          rows: rows.map((row) => ({ moduleKey: row.moduleKey, dataJson: row.dataJson, sortOrder: row.sortOrder })),
        });
        fileName = `${safeFileStem(project.projectCode || project.name)}-project-workbook.xlsx`;
      }
      const key = `exports/${fileName}`;
      await ctx.blobs.put(key, bytes, { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = await ctx.blobs.getUrl(key);
      return { ok: true, url, fileName, message: "Workbook ready." };
    },
  }),

  getPortfolioDashboard: defineAction({
    request: z.object({ asOfDate: z.string().regex(/^\d{4}-\d{2}-\d{2}$/) }),
    response: portfolioResponse,
    async handler(ctx, args): Promise<z.infer<typeof portfolioResponse>> {
      const db = ctx.db<typeof schema>();
      const [projectRows, allRows, allSections] = await Promise.all([
        db.select().from(schema.projects).orderBy(desc(schema.projects.updatedAt)),
        db.select().from(schema.moduleRows).orderBy(asc(schema.moduleRows.sortOrder), asc(schema.moduleRows.id)),
        db.select().from(schema.sections),
      ]);
      const projects: z.infer<typeof portfolioHealthShape>[] = [];
      const budgets: z.infer<typeof portfolioBudgetItemShape>[] = [];
      const dates: z.infer<typeof portfolioDateShape>[] = [];
      const risks: z.infer<typeof portfolioRiskShape>[] = [];
      const taskStatus = { notStarted: 0, inProgress: 0, complete: 0, blocked: 0 };
      const dayMs = 86_400_000;
      const asOf = new Date(`${args.asOfDate}T00:00:00Z`);
      const horizon = new Date(asOf);
      horizon.setUTCDate(horizon.getUTCDate() + 14);
      const horizonKey = horizon.toISOString().slice(0, 10);

      for (const project of projectRows) {
        const rows = allRows.filter((row) => row.projectId === project.id);
        const sections = allSections.filter((section) => section.projectId === project.id);
        const tasks = rows.filter((row) => row.moduleKey === "task-tracker");
        const timeline = rows.filter((row) => row.moduleKey === "timeline");
        const wbs = rows.filter((row) => row.moduleKey === "wbs");
        const requirements = rows.filter((row) => row.moduleKey === "requirements");
        const raciRows = rows.filter((row) => row.moduleKey === "raci");
        const acRows = rows.filter((row) => row.moduleKey === "ac-log" || row.moduleKey === "ac_log");
        const calendarRows = rows.filter((row) => row.moduleKey === "calendar" || row.moduleKey === "events");
        const taskData = tasks.map((row) => ({ row, data: parseRowData(row.dataJson) }));
        for (const { data } of taskData) {
          const status = rowText(data, "status").toLowerCase();
          if (status === "blocked") taskStatus.blocked += 1;
          else if (["complete", "completed", "done", "closed"].includes(status) || rowNumber(data, "percentComplete", "percent_complete") >= 100) taskStatus.complete += 1;
          else if (status === "in progress" || rowNumber(data, "percentComplete", "percent_complete") > 0) taskStatus.inProgress += 1;
          else taskStatus.notStarted += 1;
        }
        const wbsData = wbs.map((row) => ({ row, data: parseRowData(row.dataJson) }));
        const timelineData = timeline.map((row) => ({ row, data: parseRowData(row.dataJson) }));
        const percentComplete = taskData.length ? Math.round(taskData.reduce((sum, item) => sum + Math.max(0, Math.min(100, rowNumber(item.data, "percentComplete", "percent_complete"))), 0) / taskData.length) : 0;

        const overdueTasks = taskData.filter(({ data }) => !isCompleteRow(data) && Boolean(rowText(data, "dueDate", "due_date", "due")) && rowText(data, "dueDate", "due_date", "due") < args.asOfDate);
        const unownedTasks = taskData.filter(({ data }) => !isCompleteRow(data) && !rowText(data, "owner", "assignee"));
        const atRiskTasks = taskData.filter(({ data }) => !isCompleteRow(data) && (rowText(data, "status").toLowerCase() === "blocked" || rowText(data, "status").toLowerCase() === "at risk" || hasMeaningfulBlocker(data)));
        const costPackages = wbsData.filter(({ data }) => rowText(data, "budgetedCost", "budgeted_cost") !== "" || rowText(data, "actualCost", "actual_cost") !== "");
        const budgeted = costPackages.reduce((sum, item) => sum + rowNumber(item.data, "budgetedCost", "budgeted_cost"), 0);
        const actual = costPackages.reduce((sum, item) => sum + rowNumber(item.data, "actualCost", "actual_cost"), 0);
        const overBudget = costPackages.filter(({ data }) => rowNumber(data, "actualCost", "actual_cost") > rowNumber(data, "budgetedCost", "budgeted_cost"));
        const highImpactConstraints = acRows.filter((row) => {
          const data = parseRowData(row.dataJson);
          const type = rowText(data, "type").toLowerCase();
          const impact = rowText(data, "impact").toLowerCase();
          const status = rowText(data, "status").toLowerCase();
          return type === "constraint" && ["high", "h"].includes(impact) && !["closed", "resolved", "complete", "completed"].includes(status);
        });

        const reasons: string[] = [];
        if (overdueTasks.length) reasons.push(`${overdueTasks.length} overdue incomplete ${overdueTasks.length === 1 ? "task" : "tasks"}`);
        if (overBudget.length) reasons.push(`${overBudget.length} over-budget work ${overBudget.length === 1 ? "package" : "packages"}`);
        if (highImpactConstraints.length) reasons.push(`${highImpactConstraints.length} open high-impact ${highImpactConstraints.length === 1 ? "constraint" : "constraints"}`);
        const red = reasons.length > 0;
        if (!red && atRiskTasks.length) reasons.push(`${atRiskTasks.length} at-risk ${atRiskTasks.length === 1 ? "item" : "items"}`);
        if (!red && unownedTasks.length) reasons.push(`${unownedTasks.length} unowned incomplete ${unownedTasks.length === 1 ? "task" : "tasks"}`);
        const health: "Red" | "Amber" | "Green" = red ? "Red" : reasons.length ? "Amber" : "Green";
        if (!reasons.length) reasons.push("No red or amber conditions found");

        const linkedRequirementIds = new Set(wbsData.flatMap(({ data }) => splitReferences(rowText(data, "requirementIds", "requirement_ids", "linkedRequirementIds"))));
        const uncoveredRequirements = requirements.filter((row) => {
          const id = rowText(parseRowData(row.dataJson), "requirementId", "requirement_id", "req_id").toLowerCase();
          return !id || !linkedRequirementIds.has(id);
        }).length;
        const promotedWbsIds = new Set(tasks.map((row) => row.sourceRowId).filter((id): id is number => id !== null));
        const unpromotedWbs = wbs.filter((row) => !promotedWbsIds.has(row.id)).length;
        const traceability = uncoveredRequirements + unpromotedWbs + unownedTasks.length;

        const raciSection = sections.find((section) => section.moduleKey === "raci");
        const roles = splitReferences(rowText(parseRowData(raciSection?.dataJson ?? "{}"), "roles"));
        const raci = raciRows.reduce((count, row) => {
          const data = parseRowData(row.dataJson);
          const accountable = roles.reduce((sum, _role, index) => sum + (rowText(data, `role${index}`).toUpperCase() === "A" ? 1 : 0), 0);
          return count + (accountable === 1 ? 0 : 1);
        }, 0);

        let schedule = 0;
        const threeDays = new Date(asOf);
        threeDays.setUTCDate(threeDays.getUTCDate() + 3);
        const threeDaysKey = threeDays.toISOString().slice(0, 10);
        for (const { data } of taskData) {
          const due = rowText(data, "dueDate", "due_date", "due");
          const start = rowText(data, "startDate", "start_date", "start");
          if (!isCompleteRow(data) && due && due < args.asOfDate) schedule += 1;
          if (!isCompleteRow(data) && rowText(data, "status").toLowerCase() === "not started" && start && start >= args.asOfDate && start <= threeDaysKey) schedule += 1;
          if (!isCompleteRow(data) && !rowText(data, "owner", "assignee")) schedule += 1;
        }
        for (const { row, data } of timelineData) {
          const end = rowText(data, "endDate", "end_date", "end");
          const start = rowText(data, "startDate", "start_date", "start");
          if (row.sourceRowId !== null && !isCompleteRow(data) && ((end && end < args.asOfDate) || (!end && start && start < args.asOfDate))) schedule += 1;
          if (rowText(data, "milestone").toLowerCase() === "yes" && !rowText(data, "owner")) schedule += 1;
        }

        const targetCandidates = [
          ...taskData.map(({ data }) => rowText(data, "dueDate", "due_date", "due")),
          ...timelineData.map(({ data }) => rowText(data, "endDate", "end_date", "end")),
          ...wbsData.map(({ data }) => rowText(data, "end", "endDate", "end_date")),
        ].filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(value));
        const targetDate = targetCandidates.sort().at(-1) ?? "";
        const daysToTarget = targetDate ? Math.round((new Date(`${targetDate}T00:00:00Z`).getTime() - asOf.getTime()) / dayMs) : null;

        budgets.push({ projectId: project.id, projectName: project.name, projectCode: project.projectCode, budgeted, actual, variance: actual - budgeted, hasBudgetData: costPackages.length > 0 });
        projects.push({ projectId: project.id, projectName: project.name, projectCode: project.projectCode, health, reasons, percentComplete, taskCount: tasks.length, targetDate, daysToTarget, budgeted, actual, variance: actual - budgeted, attention: { traceability, raci, schedule } });

        for (const { row, data } of taskData) {
          if (isCompleteRow(data)) continue;
          const due = rowText(data, "dueDate", "due_date", "due");
          if (due && (due < args.asOfDate || due <= horizonKey)) dates.push({ projectId: project.id, projectName: project.name, projectCode: project.projectCode, title: rowText(data, "task", "taskName") || `Task ${row.id}`, date: due, type: "Task", owner: rowText(data, "owner") || "Unassigned", overdue: due < args.asOfDate });
          if (rowText(data, "status").toLowerCase() === "blocked" || hasMeaningfulBlocker(data)) risks.push({ projectId: project.id, projectName: project.name, projectCode: project.projectCode, kind: "Blocker", title: rowText(data, "task", "taskName") || `Task ${row.id}`, detail: rowText(data, "blocker", "risk", "blockerRisk") || "Marked blocked; details not entered.", owner: rowText(data, "owner") || "Unassigned", impact: rowText(data, "priority") || "Not rated" });
        }
        for (const { row, data } of timelineData) {
          if (isCompleteRow(data) || rowText(data, "milestone").toLowerCase() !== "yes") continue;
          const date = rowText(data, "endDate", "end_date", "end", "startDate", "start_date", "start");
          if (date && (date < args.asOfDate || date <= horizonKey)) dates.push({ projectId: project.id, projectName: project.name, projectCode: project.projectCode, title: rowText(data, "taskName", "task", "deliverable") || `Milestone ${row.id}`, date, type: "Milestone", owner: rowText(data, "owner") || "Unassigned", overdue: date < args.asOfDate });
        }
        for (const row of calendarRows) {
          const data = parseRowData(row.dataJson);
          const date = rowText(data, "date");
          if (date && date >= args.asOfDate && date <= horizonKey) dates.push({ projectId: project.id, projectName: project.name, projectCode: project.projectCode, title: rowText(data, "title", "item", "description") || `Key date ${row.id}`, date, type: "Key date", owner: rowText(data, "owner", "organizer") || "Unassigned", overdue: false });
        }
        for (const { row, data } of wbsData) {
          if (promotedWbsIds.has(row.id)) continue;
          const date = rowText(data, "end", "endDate", "end_date");
          if (date && (date < args.asOfDate || date <= horizonKey)) dates.push({ projectId: project.id, projectName: project.name, projectCode: project.projectCode, title: rowText(data, "deliverable") || `Work package ${row.id}`, date, type: "Work package", owner: rowText(data, "owner") || "Unassigned", overdue: date < args.asOfDate });
        }
        for (const row of acRows) {
          const data = parseRowData(row.dataJson);
          const kindValue = rowText(data, "type").toLowerCase();
          const impactValue = rowText(data, "impact").toLowerCase();
          const status = rowText(data, "status").toLowerCase();
          if (!["high", "h"].includes(impactValue) || ["closed", "resolved", "complete", "completed"].includes(status)) continue;
          const kind: "Assumption" | "Constraint" = kindValue === "constraint" ? "Constraint" : "Assumption";
          risks.push({ projectId: project.id, projectName: project.name, projectCode: project.projectCode, kind, title: rowText(data, "itemId", "item_id", "ac_id") || `${kind} ${row.id}`, detail: rowText(data, "description") || "No description entered.", owner: rowText(data, "owner") || "Unassigned", impact: "High" });
        }
      }

      const healthOrder = { Red: 0, Amber: 1, Green: 2 } as const;
      projects.sort((a, b) => healthOrder[a.health] - healthOrder[b.health] || a.projectName.localeCompare(b.projectName));
      dates.sort((a, b) => Number(b.overdue) - Number(a.overdue) || a.date.localeCompare(b.date) || a.projectName.localeCompare(b.projectName));
      risks.sort((a, b) => (a.kind === "Blocker" ? 0 : 1) - (b.kind === "Blocker" ? 0 : 1) || a.projectName.localeCompare(b.projectName));
      const totalBudgeted = budgets.reduce((sum, item) => sum + item.budgeted, 0);
      const totalActual = budgets.reduce((sum, item) => sum + item.actual, 0);
      return {
        ok: true,
        refreshedAt: new Date().toISOString(),
        asOfDate: args.asOfDate,
        totals: { projects: projects.length, red: projects.filter((item) => item.health === "Red").length, amber: projects.filter((item) => item.health === "Amber").length, green: projects.filter((item) => item.health === "Green").length, budgeted: totalBudgeted, actual: totalActual, variance: totalActual - totalBudgeted, taskStatus },
        projects, budgets, dates, risks,
      };
    },
  }),

  createProject: defineAction({
    request: z.object({
      projectCode: z.string().trim().min(1).max(40), name: z.string().trim().min(1).max(120),
      sponsor: z.string().max(120).default(""), projectManager: z.string().max(120).default(""),
      businessOwner: z.string().max(120).default(""), projectDate: z.string().default(""),
      version: z.string().max(30).default("1.0"), overview: z.string().max(4000).default(""),
    }),
    response: z.object({ ok: z.boolean(), id: z.number().optional(), message: z.string() }),
    async handler(ctx, args) {
      const db = ctx.db<typeof schema>();
      const existing = await db.select({ id: schema.projects.id }).from(schema.projects).where(eq(schema.projects.projectCode, args.projectCode)).limit(1);
      if (existing.length) return { ok: false, message: "That Project ID is already in use." };
      const inserted = await db.insert(schema.projects).values(args).returning({ id: schema.projects.id });
      const row = inserted[0];
      if (!row) return { ok: false, message: "The project could not be created." };
      ctx.invalidateQueries();
      return { ok: true, id: row.id, message: "Project created." };
    },
  }),

  importProjectWorkbook: defineAction({
    request: z.object({
      project: workbookProjectShape,
      sections: z.array(workbookSectionShape).max(20),
      rows: z.array(workbookRowShape).max(5000),
    }),
    response: importWorkbookResponse,
    async handler(ctx, args): Promise<z.infer<typeof importWorkbookResponse>> {
      const db = ctx.db<typeof schema>();
      const cleanSections = args.sections.flatMap((item) => {
        const dataJson = validatedWorkbookJson(item.dataJson);
        return dataJson && workbookSectionKeys.has(item.moduleKey) ? [{ ...item, dataJson }] : [];
      });
      const cleanRows = args.rows.flatMap((item) => {
        const dataJson = validatedWorkbookJson(item.dataJson);
        return dataJson && workbookRowKeys.has(item.moduleKey) ? [{ ...item, dataJson }] : [];
      });
      let projectCode = args.project.projectCode.trim();
      let suffix = 0;
      while (true) {
        const match = await db.select({ id: schema.projects.id }).from(schema.projects).where(eq(schema.projects.projectCode, projectCode)).limit(1);
        if (!match.length) break;
        suffix += 1;
        const ending = suffix === 1 ? "-COPY" : `-COPY-${suffix}`;
        projectCode = `${args.project.projectCode.slice(0, Math.max(1, 40 - ending.length))}${ending}`;
      }
      const inserted = await db.insert(schema.projects).values({ ...args.project, projectCode }).returning({ id: schema.projects.id });
      const project = inserted[0];
      if (!project) return { ok: false, message: "The imported project could not be created.", sectionsCreated: 0, rowsCreated: 0 };
      try {
        if (cleanSections.length) await db.insert(schema.sections).values(cleanSections.map((item) => ({ projectId: project.id, moduleKey: item.moduleKey, dataJson: item.dataJson })));
        if (cleanRows.length) await db.insert(schema.moduleRows).values(cleanRows.map((item) => ({ projectId: project.id, moduleKey: item.moduleKey, dataJson: item.dataJson, sortOrder: item.sortOrder })));
      } catch {
        await db.delete(schema.projects).where(eq(schema.projects.id, project.id));
        return { ok: false, message: "The workbook was valid, but its project records could not be saved.", sectionsCreated: 0, rowsCreated: 0 };
      }
      ctx.invalidateQueries();
      const renamed = projectCode !== args.project.projectCode;
      return {
        ok: true,
        id: project.id,
        projectCode,
        message: renamed ? `Project created as ${projectCode} because the original Project ID was already in use.` : "Project created from the workbook.",
        sectionsCreated: cleanSections.length,
        rowsCreated: cleanRows.length,
      };
    },
  }),

  updateProject: defineAction({
    request: z.object({
      id: z.number().int().positive(), projectCode: z.string().trim().min(1).max(40), name: z.string().trim().min(1).max(120),
      sponsor: z.string().max(120), projectManager: z.string().max(120), businessOwner: z.string().max(120),
      projectDate: z.string(), version: z.string().max(30), overview: z.string().max(4000),
    }),
    response: mutationResponse,
    async handler(ctx, args) {
      const db = ctx.db<typeof schema>();
      const duplicate = await db.select({ id: schema.projects.id }).from(schema.projects).where(eq(schema.projects.projectCode, args.projectCode)).limit(1);
      if (duplicate[0] && duplicate[0].id !== args.id) return { ok: false, message: "That Project ID is already in use." };
      const { id, ...values } = args;
      await db.update(schema.projects).set({ ...values, updatedAt: new Date() }).where(eq(schema.projects.id, id));
      ctx.invalidateQueries();
      return { ok: true, message: "Project header saved." };
    },
  }),

  getProjectData: defineAction({
    request: z.object({ projectId: z.number().int().positive() }),
    response: z.object({
      project: projectShape.optional(),
      sections: z.array(z.object({ moduleKey: z.string(), dataJson: z.string(), updatedAt: z.string() })),
      rows: z.array(z.object({ id: z.number(), moduleKey: z.string(), dataJson: z.string(), sortOrder: z.number(), sourceRowId: z.number().nullable(), updatedAt: z.string() })),
    }),
    async handler(ctx, args) {
      const db = ctx.db<typeof schema>();
      const [projects, sections, rows] = await Promise.all([
        db.select().from(schema.projects).where(eq(schema.projects.id, args.projectId)).limit(1),
        db.select().from(schema.sections).where(eq(schema.sections.projectId, args.projectId)),
        db.select().from(schema.moduleRows).where(eq(schema.moduleRows.projectId, args.projectId)).orderBy(asc(schema.moduleRows.sortOrder), asc(schema.moduleRows.id)),
      ]);
      const p = projects[0];
      return {
        project: p ? { ...p, updatedAt: p.updatedAt.toISOString() } : undefined,
        sections: sections.map((s) => ({ moduleKey: s.moduleKey, dataJson: s.dataJson, updatedAt: s.updatedAt.toISOString() })),
        rows: rows.map((r) => ({ id: r.id, moduleKey: r.moduleKey, dataJson: r.dataJson, sortOrder: r.sortOrder, sourceRowId: r.sourceRowId, updatedAt: r.updatedAt.toISOString() })),
      };
    },
  }),

  saveSection: defineAction({
    request: z.object({ projectId: z.number().int().positive(), moduleKey: z.string().min(1).max(60), dataJson: z.string().max(50000) }),
    response: mutationResponse,
    async handler(ctx, args) {
      try { JSON.parse(args.dataJson); } catch { return { ok: false, message: "This section contains invalid data." }; }
      const db = ctx.db<typeof schema>();
      const existing = await db.select({ id: schema.sections.id }).from(schema.sections).where(and(eq(schema.sections.projectId, args.projectId), eq(schema.sections.moduleKey, args.moduleKey))).limit(1);
      if (existing[0]) await db.update(schema.sections).set({ dataJson: args.dataJson, updatedAt: new Date() }).where(eq(schema.sections.id, existing[0].id));
      else await db.insert(schema.sections).values({ ...args });
      await db.update(schema.projects).set({ updatedAt: new Date() }).where(eq(schema.projects.id, args.projectId));
      ctx.invalidateQueries();
      return { ok: true, message: "Section saved." };
    },
  }),

  addRow: defineAction({
    request: z.object({ projectId: z.number().int().positive(), moduleKey: z.string().min(1).max(60), dataJson: z.string().max(50000) }),
    response: z.object({ ok: z.boolean(), id: z.number().optional(), message: z.string() }),
    async handler(ctx, args) {
      let rowData: RowData;
      try { rowData = JSON.parse(args.dataJson) as RowData; } catch { return { ok: false, message: "This row contains invalid data." }; }
      if (args.moduleKey === "task-tracker" && rowData.status === "Complete") rowData.completedAt = new Date().toISOString();
      const db = ctx.db<typeof schema>();
      const existing = await db.select({ sortOrder: schema.moduleRows.sortOrder }).from(schema.moduleRows).where(and(eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, args.moduleKey))).orderBy(desc(schema.moduleRows.sortOrder)).limit(1);
      const inserted = await db.insert(schema.moduleRows).values({ ...args, dataJson: JSON.stringify(rowData), sortOrder: (existing[0]?.sortOrder ?? -1) + 1 }).returning({ id: schema.moduleRows.id });
      ctx.invalidateQueries();
      return { ok: true, id: inserted[0]?.id, message: "Row added." };
    },
  }),

  updateRow: defineAction({
    request: z.object({ id: z.number().int().positive(), projectId: z.number().int().positive(), dataJson: z.string().max(50000) }),
    response: mutationResponse,
    async handler(ctx, args) {
      let nextData: RowData;
      try { nextData = JSON.parse(args.dataJson) as RowData; } catch { return { ok: false, message: "This row contains invalid data." }; }
      const db = ctx.db<typeof schema>();
      const existing = await db.select({ moduleKey: schema.moduleRows.moduleKey, dataJson: schema.moduleRows.dataJson }).from(schema.moduleRows).where(and(eq(schema.moduleRows.id, args.id), eq(schema.moduleRows.projectId, args.projectId))).limit(1);
      const current = existing[0];
      if (!current) return { ok: false, message: "Row not found." };
      if (current.moduleKey === "task-tracker") {
        const previousData = parseRowData(current.dataJson);
        if (nextData.status === "Complete" && previousData.status !== "Complete") nextData.completedAt = new Date().toISOString();
        if (nextData.status !== "Complete") delete nextData.completedAt;
      }
      await db.update(schema.moduleRows).set({ dataJson: JSON.stringify(nextData), updatedAt: new Date() }).where(and(eq(schema.moduleRows.id, args.id), eq(schema.moduleRows.projectId, args.projectId)));
      ctx.invalidateQueries();
      return { ok: true, message: "Row saved." };
    },
  }),

  deleteRow: defineAction({
    request: z.object({ id: z.number().int().positive(), projectId: z.number().int().positive() }),
    response: mutationResponse,
    async handler(ctx, args) {
      const db = ctx.db<typeof schema>();
      await db.delete(schema.moduleRows).where(and(eq(schema.moduleRows.id, args.id), eq(schema.moduleRows.projectId, args.projectId)));
      ctx.invalidateQueries();
      return { ok: true, message: "Row removed." };
    },
  }),

  promoteWbs: defineAction({
    request: z.object({ projectId: z.number().int().positive(), rowId: z.number().int().positive() }),
    response: mutationResponse,
    async handler(ctx, args) {
      const db = ctx.db<typeof schema>();
      const source = await db.select().from(schema.moduleRows).where(and(eq(schema.moduleRows.id, args.rowId), eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, "wbs"))).limit(1);
      const item = source[0];
      if (!item) return { ok: false, message: "WBS item not found." };
      const existing = await db.select({ id: schema.moduleRows.id }).from(schema.moduleRows).where(and(eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, "task-tracker"), eq(schema.moduleRows.sourceRowId, args.rowId))).limit(1);
      if (existing.length) return { ok: false, message: "This WBS item is already in the task tracker." };
      const data = JSON.parse(item.dataJson) as Record<string, string>;
      const task = { taskId: data.wbsId ?? "", task: data.deliverable ?? "", owner: data.owner ?? "", startDate: data.start ?? "", dueDate: data.end ?? "", priority: "Medium", status: "Not Started", percentComplete: "0", blocker: "", nextStep: "", dependency: data.dependencies ?? "", notes: data.description ?? "" };
      await db.insert(schema.moduleRows).values({ projectId: args.projectId, moduleKey: "task-tracker", dataJson: JSON.stringify(task), sortOrder: item.sortOrder, sourceRowId: args.rowId });
      ctx.invalidateQueries();
      return { ok: true, message: "WBS item promoted to the task tracker." };
    },
  }),

  spawnKickoffTask: defineAction({
    request: z.object({ projectId: z.number().int().positive(), rowId: z.number().int().positive(), owner: z.string().max(120), dueDate: z.string().regex(/^$|^\d{4}-\d{2}-\d{2}$/) }),
    response: mutationResponse,
    async handler(ctx, args) {
      const db = ctx.db<typeof schema>();
      const source = await db.select().from(schema.moduleRows).where(and(eq(schema.moduleRows.id, args.rowId), eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, "kickoff"))).limit(1);
      const item = source[0];
      if (!item) return { ok: false, message: "Kickoff item not found." };
      const existing = await db.select({ id: schema.moduleRows.id }).from(schema.moduleRows).where(and(eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, "task-tracker"), eq(schema.moduleRows.sourceRowId, args.rowId))).limit(1);
      if (existing.length) return { ok: false, message: "This kickoff item is already linked to a tracked task." };
      const data = parseRowData(item.dataJson);
      const itemText = data.item?.trim();
      if (!itemText) return { ok: false, message: "Add item text before creating a task." };
      const latest = await db.select({ sortOrder: schema.moduleRows.sortOrder }).from(schema.moduleRows).where(and(eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, "task-tracker"))).orderBy(desc(schema.moduleRows.sortOrder)).limit(1);
      const task = { taskId: `KICK-${item.id}`, task: itemText, owner: args.owner.trim(), startDate: "", dueDate: args.dueDate, priority: "Medium", status: "Not Started", percentComplete: "0", blocker: "", nextStep: "", dependency: "", notes: data.notes?.trim() ? `Kickoff note: ${data.notes.trim()}` : "Created from kickoff agenda." };
      await db.insert(schema.moduleRows).values({ projectId: args.projectId, moduleKey: "task-tracker", dataJson: JSON.stringify(task), sortOrder: (latest[0]?.sortOrder ?? -1) + 1, sourceRowId: args.rowId });
      ctx.invalidateQueries();
      return { ok: true, message: "Kickoff item linked to a tracked task." };
    },
  }),

  sendPertToTimeline: defineAction({
    request: z.object({ projectId: z.number().int().positive(), rowId: z.number().int().positive() }),
    response: mutationResponse,
    async handler(ctx, args) {
      const db = ctx.db<typeof schema>();
      const source = await db.select().from(schema.moduleRows).where(and(eq(schema.moduleRows.id, args.rowId), eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, "pert"))).limit(1);
      const item = source[0];
      if (!item) return { ok: false, message: "PERT estimate not found." };
      const existing = await db.select({ id: schema.moduleRows.id }).from(schema.moduleRows).where(and(eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, "timeline"), eq(schema.moduleRows.sourceRowId, args.rowId))).limit(1);
      if (existing.length) return { ok: false, message: "This estimate is already on the timeline." };
      const data = JSON.parse(item.dataJson) as Record<string, string>;
      const o = Number(data.optimistic || 0), m = Number(data.mostLikely || 0), p = Number(data.pessimistic || 0);
      const expected = (o + 4 * m + p) / 6;
      const timeline = { wbsId: data.activityId ?? "", taskName: data.activityName ?? "", owner: data.owner ?? "", startDate: "", endDate: "", duration: expected.toFixed(2), percentComplete: "0", status: "Not Started", predecessors: data.predecessors ?? "", dependencyType: "FS", milestone: "No", notes: `PERT expected duration; σ=${((p - o) / 6).toFixed(2)}` };
      await db.insert(schema.moduleRows).values({ projectId: args.projectId, moduleKey: "timeline", dataJson: JSON.stringify(timeline), sortOrder: item.sortOrder, sourceRowId: args.rowId });
      ctx.invalidateQueries();
      return { ok: true, message: "PERT duration added to the timeline." };
    },
  }),

  runTraceabilityCheck: defineAction({
    request: z.object({ projectId: z.number().int().positive() }),
    response: traceabilityResponse,
    async handler(ctx, args): Promise<z.infer<typeof traceabilityResponse>> {
      const db = ctx.db<typeof schema>();
      const rows = await db.select().from(schema.moduleRows).where(eq(schema.moduleRows.projectId, args.projectId));
      const requirements = rows.filter((row) => row.moduleKey === "requirements");
      const wbs = rows.filter((row) => row.moduleKey === "wbs");
      const tasks = rows.filter((row) => row.moduleKey === "task-tracker");
      const linkedRequirementIds = new Set(wbs.flatMap((row) => splitReferences(parseRowData(row.dataJson).requirementIds)));
      const promotedWbsIds = new Set(tasks.map((row) => row.sourceRowId).filter((id): id is number => id !== null));
      const gaps: z.infer<typeof gapShape>[] = [];

      for (const row of requirements) {
        const data = parseRowData(row.dataJson);
        const requirementId = data.requirementId?.trim();
        if (!requirementId || !linkedRequirementIds.has(requirementId.toLowerCase())) {
          gaps.push({ type: "requirement", label: requirementId || displayTask(data, `Requirement ${row.id}`), detail: "No WBS work package links to this requirement.", moduleKey: "requirements", rowId: row.id });
        }
      }
      for (const row of wbs) {
        if (!promotedWbsIds.has(row.id)) {
          const data = parseRowData(row.dataJson);
          gaps.push({ type: "wbs", label: data.wbsId?.trim() || displayTask(data, `WBS item ${row.id}`), detail: "This WBS item has not been promoted to the task tracker.", moduleKey: "wbs", rowId: row.id });
        }
      }
      for (const row of tasks) {
        const data = parseRowData(row.dataJson);
        if (!data.owner?.trim()) gaps.push({ type: "task", label: data.taskId?.trim() || displayTask(data, `Task ${row.id}`), detail: "No owner is assigned to this task.", moduleKey: "task-tracker", rowId: row.id });
      }
      return {
        ok: true,
        message: gaps.length ? `${gaps.length} attention ${gaps.length === 1 ? "item" : "items"} found.` : "No traceability gaps found.",
        checkedAt: new Date().toISOString(),
        requirementsChecked: requirements.length,
        wbsChecked: wbs.length,
        tasksChecked: tasks.length,
        gaps,
      };
    },
  }),

  runRaciValidation: defineAction({
    request: z.object({ projectId: z.number().int().positive() }),
    response: raciResponse,
    async handler(ctx, args): Promise<z.infer<typeof raciResponse>> {
      const db = ctx.db<typeof schema>();
      const sectionRows = await db.select().from(schema.sections).where(and(eq(schema.sections.projectId, args.projectId), eq(schema.sections.moduleKey, "raci"))).limit(1);
      const roles = splitReferences(parseRowData(sectionRows[0]?.dataJson ?? "{}").roles);
      const rows = await db.select().from(schema.moduleRows).where(and(eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, "raci"))).orderBy(asc(schema.moduleRows.sortOrder), asc(schema.moduleRows.id));
      const issues: z.infer<typeof raciIssueShape>[] = [];
      for (const row of rows) {
        const data = parseRowData(row.dataJson);
        const accountableCount = roles.reduce((count, _role, index) => count + (data[`role${index}`] === "A" ? 1 : 0), 0);
        if (accountableCount !== 1) {
          issues.push({ rowId: row.id, activity: data.activity?.trim() || `Activity ${row.id}`, issue: accountableCount === 0 ? "no one accountable" : "conflicting accountability", accountableCount });
        }
      }
      return {
        ok: true,
        message: !rows.length ? "No RACI activities to validate." : issues.length ? `${issues.length} accountability ${issues.length === 1 ? "issue" : "issues"} found.` : "Every RACI activity has exactly one Accountable.",
        checkedAt: new Date().toISOString(), activitiesChecked: rows.length, issues,
      };
    },
  }),

  runScheduleHealthScan: defineAction({
    request: z.object({ projectId: z.number().int().positive(), scanDate: z.string().regex(/^\d{4}-\d{2}-\d{2}$/) }),
    response: scheduleHealthResponse,
    async handler(ctx, args): Promise<z.infer<typeof scheduleHealthResponse>> {
      const db = ctx.db<typeof schema>();
      const rows = await db.select().from(schema.moduleRows).where(eq(schema.moduleRows.projectId, args.projectId)).orderBy(asc(schema.moduleRows.sortOrder), asc(schema.moduleRows.id));
      const tasks = rows.filter((row) => row.moduleKey === "task-tracker");
      const timelineItems = rows.filter((row) => row.moduleKey === "timeline");
      const risks: z.infer<typeof scheduleRiskShape>[] = [];
      const scanStart = new Date(`${args.scanDate}T00:00:00Z`);
      const threeDaysOut = new Date(scanStart);
      threeDaysOut.setUTCDate(threeDaysOut.getUTCDate() + 3);
      const threeDaysKey = threeDaysOut.toISOString().slice(0, 10);

      for (const row of tasks) {
        const data = parseRowData(row.dataJson);
        const taskName = displayTask(data, `Task ${row.id}`);
        if (data.status !== "Complete" && data.dueDate && data.dueDate < args.scanDate) risks.push({ rowId: row.id, moduleKey: "task-tracker", taskName, reason: "Past due and not complete", relevantDate: data.dueDate });
        if (data.status === "Not Started" && data.startDate && data.startDate >= args.scanDate && data.startDate <= threeDaysKey) risks.push({ rowId: row.id, moduleKey: "task-tracker", taskName, reason: "Starts within 3 days and is still Not Started", relevantDate: data.startDate });
        if (!data.owner?.trim()) risks.push({ rowId: row.id, moduleKey: "task-tracker", taskName, reason: "Task has no owner", relevantDate: data.dueDate || data.startDate || "No date" });
      }
      for (const row of timelineItems) {
        const data = parseRowData(row.dataJson);
        const taskName = displayTask(data, `Timeline item ${row.id}`);
        if (row.sourceRowId !== null && data.status !== "Complete") {
          const passedDate = data.endDate && data.endDate < args.scanDate ? data.endDate : !data.endDate && data.startDate && data.startDate < args.scanDate ? data.startDate : "";
          if (passedDate) risks.push({ rowId: row.id, moduleKey: "timeline", taskName, reason: "PERT-derived timeline date passed without completion", relevantDate: passedDate });
        }
        if (data.milestone === "Yes" && !data.owner?.trim()) risks.push({ rowId: row.id, moduleKey: "timeline", taskName, reason: "Milestone has no owner", relevantDate: data.endDate || data.startDate || "No date" });
      }
      return {
        ok: true,
        message: !tasks.length && !timelineItems.length ? "No tasks or timeline items to scan." : risks.length ? `${risks.length} schedule ${risks.length === 1 ? "risk" : "risks"} found.` : "No schedule risks found.",
        checkedAt: new Date().toISOString(), tasksChecked: tasks.length, timelineItemsChecked: timelineItems.length, risks,
      };
    },
  }),

  compileStatusReport: defineAction({
    request: z.object({ projectId: z.number().int().positive(), reportDate: z.string().regex(/^\d{4}-\d{2}-\d{2}$/) }),
    response: reportResponse,
    async handler(ctx, args): Promise<z.infer<typeof reportResponse>> {
      const db = ctx.db<typeof schema>();
      const projectRows = await db.select().from(schema.projects).where(eq(schema.projects.id, args.projectId)).limit(1);
      const project = projectRows[0];
      if (!project) return { ok: true, message: "Project not found.", draft: "" };
      const rows = await db.select().from(schema.moduleRows).where(and(eq(schema.moduleRows.projectId, args.projectId), eq(schema.moduleRows.moduleKey, "task-tracker"))).orderBy(asc(schema.moduleRows.sortOrder), asc(schema.moduleRows.id));
      const tasks = rows.map((row) => ({ row, data: parseRowData(row.dataJson) }));
      const reportEnd = new Date(`${args.reportDate}T23:59:59.999Z`);
      const reportStart = new Date(reportEnd);
      reportStart.setUTCDate(reportEnd.getUTCDate() - 6);
      reportStart.setUTCHours(0, 0, 0, 0);
      const completed = tasks.filter(({ data }) => {
        if (data.status !== "Complete") return false;
        if (!data.completedAt) return false;
        const completedAt = new Date(data.completedAt);
        return completedAt >= reportStart && completedAt <= reportEnd;
      });
      const atRisk = tasks.filter(({ data }) => data.status !== "Complete" && (data.status === "Blocked" || Boolean(data.blocker?.trim()) || Boolean(data.dueDate && data.dueDate < args.reportDate)));
      const blockers = tasks.filter(({ data }) => data.status !== "Complete" && (data.status === "Blocked" || Boolean(data.blocker?.trim())));
      const overall = tasks.length ? Math.round(tasks.reduce((sum, { data }) => sum + Math.max(0, Math.min(100, Number(data.percentComplete || 0) || 0)), 0) / tasks.length) : 0;
      const line = (data: RowData, fallback: string) => `- ${displayTask(data, fallback)}${data.owner?.trim() ? `, owner: ${data.owner.trim()}` : ", owner: Unassigned"}`;
      const section = (title: string, items: string[], empty: string) => `${title}\n${items.length ? items.join("\n") : `- ${empty}`}`;
      const draft = [
        `WEEKLY PROJECT STATUS REPORT`,
        `Project: ${project.name}`,
        `Project Code: ${project.projectCode}`,
        `Report Date: ${args.reportDate}`,
        `Reporting Window: ${reportStart.toISOString().slice(0, 10)} to ${args.reportDate}`,
        "",
        `OVERALL PROGRESS\n${overall}% complete across ${tasks.length} ${tasks.length === 1 ? "task" : "tasks"}.`,
        "",
        section("COMPLETED THIS WEEK", completed.map(({ row, data }) => line(data, `Task ${row.id}`)), "No tasks were completed this week."),
        "",
        section("TASKS AT RISK", atRisk.map(({ row, data }) => `${line(data, `Task ${row.id}`)}${data.status === "Blocked" || data.blocker?.trim() ? ` [Blocked${data.blocker?.trim() ? `: ${data.blocker.trim()}` : ""}]` : ` [Past due: ${data.dueDate}]`}`), "No past-due or blocked tasks."),
        "",
        section("OPEN BLOCKERS", blockers.map(({ row, data }) => `${line(data, `Task ${row.id}`)}: ${data.blocker?.trim() || "Marked blocked; blocker details not entered."}`), "No open blockers."),
        "",
        "NOTES / NEXT STEPS\nAdd updates, decisions, and next steps here before sending.",
      ].join("\n");
      return { ok: true, message: "Weekly draft compiled from the current task tracker.", draft };
    },
  }),
} satisfies ActionsModule;
