import * as XLSX from "xlsx";

export type ExportProject = {
  projectCode: string;
  name: string;
  sponsor: string;
  projectManager: string;
  businessOwner: string;
  projectDate: string;
  version: string;
  overview: string;
};

export type ExportSection = { moduleKey: string; dataJson: string };
export type ExportRow = { moduleKey: string; dataJson: string; sortOrder: number };
type WorkbookField = { key: string; label: string; kind?: "date" | "checkbox" };
type SheetSpec = { name: string; moduleKey: string; sectionFields?: WorkbookField[]; rowFields?: WorkbookField[]; includeProject?: boolean };

const projectFields: WorkbookField[] = [
  { key: "projectCode", label: "Project ID" }, { key: "name", label: "Project Name" },
  { key: "sponsor", label: "Sponsor" }, { key: "projectManager", label: "Project Manager" },
  { key: "businessOwner", label: "Business Owner" }, { key: "projectDate", label: "Project Date", kind: "date" },
  { key: "version", label: "Version" }, { key: "overview", label: "Overview" },
];

const sheets: SheetSpec[] = [
  { name: "Business Case", moduleKey: "business-case", includeProject: true, sectionFields: [
    { key: "problem", label: "Problem or Opportunity" }, { key: "options", label: "Options Considered" },
    { key: "benefits", label: "Expected Benefits" }, { key: "costs", label: "Estimated Costs" }, { key: "recommendation", label: "Recommendation" },
  ] },
  { name: "Project Charter", moduleKey: "charter", sectionFields: [
    { key: "purpose", label: "Purpose and Authority" }, { key: "successCriteria", label: "Success Criteria" }, { key: "approvals", label: "Approvals" },
  ], rowFields: [
    { key: "objectiveId", label: "Objective ID" }, { key: "objective", label: "SMART Objective" }, { key: "kpi", label: "Measure / KPI" },
    { key: "target", label: "Target" }, { key: "owner", label: "Owner" },
  ] },
  { name: "Scope Statement", moduleKey: "scope", sectionFields: [
    { key: "inScope", label: "In Scope" }, { key: "outOfScope", label: "Out of Scope" }, { key: "deliverables", label: "Deliverables" }, { key: "exclusions", label: "Exclusions" },
  ] },
  { name: "Project Mgmt Plan", moduleKey: "management-plan", sectionFields: [
    { key: "approach", label: "Delivery Approach" }, { key: "governance", label: "Governance" }, { key: "cadence", label: "Meeting and Reporting Cadence" },
  ] },
  { name: "Stakeholder Register", moduleKey: "stakeholders", rowFields: [
    { key: "name", label: "Name" }, { key: "orgRole", label: "Organization / Role" }, { key: "influence", label: "Influence" },
    { key: "interest", label: "Interest" }, { key: "currentEngagement", label: "Current Engagement" }, { key: "desiredEngagement", label: "Desired Engagement" },
  ] },
  { name: "Project Objectives", moduleKey: "objectives", rowFields: [
    { key: "objective", label: "SMART Objective" }, { key: "kpi", label: "KPI" }, { key: "target", label: "Target" }, { key: "owner", label: "Owner" },
  ] },
  { name: "High-Level Requirements", moduleKey: "requirements", rowFields: [
    { key: "requirementId", label: "Requirement ID" }, { key: "description", label: "Description" }, { key: "priority", label: "Priority" },
    { key: "owner", label: "Owner" }, { key: "status", label: "Status" },
  ] },
  { name: "A&C Log", moduleKey: "ac-log", rowFields: [
    { key: "type", label: "Type" }, { key: "itemId", label: "ID" }, { key: "description", label: "Description" }, { key: "owner", label: "Owner" },
    { key: "date", label: "Review Date", kind: "date" }, { key: "impact", label: "Impact" }, { key: "status", label: "Status" },
  ] },
  { name: "Kickoff Agenda", moduleKey: "kickoff", rowFields: [
    { key: "complete", label: "Complete", kind: "checkbox" }, { key: "item", label: "Agenda Item" }, { key: "owner", label: "Owner" }, { key: "notes", label: "Decisions / Notes" },
  ] },
  { name: "RACI Matrix", moduleKey: "raci", sectionFields: [{ key: "roles", label: "Roles (comma separated)" }], rowFields: [{ key: "activity", label: "Activity" }] },
  { name: "WBS", moduleKey: "wbs", rowFields: [
    { key: "wbsId", label: "WBS ID" }, { key: "level", label: "Level" }, { key: "deliverable", label: "Deliverable / Work Package" },
    { key: "requirementIds", label: "Linked Requirement IDs" }, { key: "description", label: "Description" }, { key: "owner", label: "Owner" },
    { key: "estimateHours", label: "Estimate Hours" }, { key: "budgetedCost", label: "Budgeted Cost" }, { key: "actualCost", label: "Actual Cost" },
    { key: "start", label: "Start", kind: "date" }, { key: "end", label: "End", kind: "date" }, { key: "dependencies", label: "Dependencies" },
  ] },
  { name: "PERT", moduleKey: "pert", rowFields: [
    { key: "activityId", label: "Activity ID" }, { key: "activityName", label: "Activity Name" }, { key: "predecessors", label: "Predecessor(s)" },
    { key: "optimistic", label: "Optimistic (O)" }, { key: "mostLikely", label: "Most Likely (M)" }, { key: "pessimistic", label: "Pessimistic (P)" },
    { key: "owner", label: "Owner" }, { key: "notes", label: "Notes" },
  ] },
  { name: "Project Timeline", moduleKey: "timeline", rowFields: [
    { key: "wbsId", label: "WBS ID" }, { key: "taskName", label: "Task Name" }, { key: "owner", label: "Owner" },
    { key: "startDate", label: "Start Date", kind: "date" }, { key: "endDate", label: "End Date", kind: "date" }, { key: "duration", label: "Duration (days)" },
    { key: "percentComplete", label: "% Complete" }, { key: "status", label: "Status" }, { key: "predecessors", label: "Predecessor(s)" },
    { key: "dependencyType", label: "Dependency Type" }, { key: "milestone", label: "Milestone" }, { key: "notes", label: "Notes" },
  ] },
  { name: "Task Tracker", moduleKey: "task-tracker", rowFields: [
    { key: "taskId", label: "Task ID" }, { key: "task", label: "Task" }, { key: "owner", label: "Owner" },
    { key: "startDate", label: "Start Date", kind: "date" }, { key: "dueDate", label: "Due Date", kind: "date" },
    { key: "priority", label: "Priority" }, { key: "status", label: "Status" }, { key: "percentComplete", label: "% Complete" },
    { key: "blocker", label: "Blocker / Risk" }, { key: "nextStep", label: "Next Step" }, { key: "dependency", label: "Dependency" }, { key: "notes", label: "Notes" },
  ] },
  { name: "To Do List", moduleKey: "todo", rowFields: [
    { key: "done", label: "Done", kind: "checkbox" }, { key: "item", label: "Item" }, { key: "priority", label: "Priority" },
    { key: "dueDate", label: "Due Date", kind: "date" }, { key: "estimatedMinutes", label: "Estimated Minutes" }, { key: "category", label: "Category" }, { key: "notes", label: "Notes" },
  ] },
  { name: "Project Calendar", moduleKey: "calendar", rowFields: [
    { key: "date", label: "Date", kind: "date" }, { key: "startTime", label: "Start Time" }, { key: "endTime", label: "End Time" },
    { key: "title", label: "Event / Meeting" }, { key: "category", label: "Category" }, { key: "owner", label: "Organizer / Owner" },
    { key: "attendees", label: "Attendees" }, { key: "location", label: "Location / Link" }, { key: "description", label: "Description" },
  ] },
];

function parseRecord(value: string | undefined): Record<string, string> {
  try {
    const parsed = value ? JSON.parse(value) as unknown : {};
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return {};
    const output: Record<string, string> = {};
    for (const [key, cell] of Object.entries(parsed)) output[key] = cell === null || cell === undefined ? "" : String(cell);
    return output;
  } catch { return {}; }
}

function displayCell(value: string, field: WorkbookField): string {
  if (field.kind === "checkbox") return value === "true" ? "Yes" : value === "false" ? "No" : value;
  return value;
}

function makeSheet(headers: WorkbookField[], records: Array<Record<string, string>>): XLSX.WorkSheet {
  const body = records.map((record) => headers.map((field) => displayCell(record[field.key] ?? "", field)));
  const worksheet = XLSX.utils.aoa_to_sheet([headers.map((field) => field.label), ...body]);
  worksheet["!cols"] = headers.map((field) => ({ wch: Math.min(42, Math.max(13, field.label.length + 2)) }));
  if (headers.length) worksheet["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, body.length), c: headers.length - 1 } }) };
  return worksheet;
}

export function safeFileStem(value: string): string {
  const stem = value.trim().replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/^-+|-+$/g, "");
  return stem || "project";
}

function workbookBytes(workbook: XLSX.WorkBook): Uint8Array {
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true }) as ArrayBuffer;
  return new Uint8Array(bytes);
}

export function createProjectWorkbookBytes(data: { project: ExportProject; sections: ExportSection[]; rows: ExportRow[] }): Uint8Array {
  const workbook = XLSX.utils.book_new();
  const sectionMap = new Map(data.sections.map((section) => [section.moduleKey, parseRecord(section.dataJson)]));
  for (const spec of sheets) {
    const section = sectionMap.get(spec.moduleKey) ?? {};
    const moduleRows = data.rows.filter((row) => row.moduleKey === spec.moduleKey).sort((a, b) => a.sortOrder - b.sortOrder).map((row) => parseRecord(row.dataJson));
    let headers = [...(spec.includeProject ? projectFields : []), ...(spec.sectionFields ?? []), ...(spec.rowFields ?? [])];
    if (spec.moduleKey === "raci") {
      const roles = (section.roles ?? "").split(",").map((role) => role.trim()).filter(Boolean).slice(0, 8);
      headers = [...(spec.sectionFields ?? []), ...(spec.rowFields ?? []), ...roles.map((role, index) => ({ key: `role${index}`, label: role }))];
    }
    const base: Record<string, string> = { ...(spec.includeProject ? data.project : {}), ...section };
    const records = spec.rowFields
      ? (moduleRows.length ? moduleRows.map((row, index) => ({ ...(index === 0 ? base : {}), ...row })) : Object.values(base).some(Boolean) ? [base] : [])
      : (Object.values(base).some(Boolean) ? [base] : []);
    XLSX.utils.book_append_sheet(workbook, makeSheet(headers, records), spec.name);
  }
  return workbookBytes(workbook);
}

export function createBlankWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new();
  for (const spec of sheets) {
    let headers = [...(spec.includeProject ? projectFields : []), ...(spec.sectionFields ?? []), ...(spec.rowFields ?? [])];
    if (spec.moduleKey === "raci") headers = [...headers, ...Array.from({ length: 8 }, (_, index) => ({ key: `role${index}`, label: `Role ${index + 1} Assignment` }))];
    XLSX.utils.book_append_sheet(workbook, makeSheet(headers, []), spec.name);
  }
  return workbookBytes(workbook);
}
