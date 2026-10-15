import * as XLSX from "xlsx";

type ProjectRecord = {
  projectCode: string;
  name: string;
  sponsor: string;
  projectManager: string;
  businessOwner: string;
  projectDate: string;
  version: string;
  overview: string;
};

type SectionRecord = { moduleKey: string; dataJson: string };
type RowRecord = { id?: number; moduleKey: string; dataJson: string; sortOrder: number };
type WorkbookField = { key: string; label: string; kind?: "date" | "checkbox" };
type SheetSpec = {
  name: string;
  moduleKey: string;
  sectionFields?: WorkbookField[];
  rowFields?: WorkbookField[];
  includeProject?: boolean;
};

export type WorkbookExportData = {
  project: ProjectRecord;
  sections: Array<SectionRecord>;
  rows: Array<RowRecord>;
};

export type WorkbookImportPreview = {
  fileName: string;
  project: ProjectRecord;
  sections: Array<SectionRecord>;
  rows: Array<Omit<RowRecord, "id">>;
  counts: Array<{ sheet: string; count: number }>;
  notes: string[];
  canImport: boolean;
};

const projectFields: WorkbookField[] = [
  { key: "projectCode", label: "Project ID" },
  { key: "name", label: "Project Name" },
  { key: "sponsor", label: "Sponsor" },
  { key: "projectManager", label: "Project Manager" },
  { key: "businessOwner", label: "Business Owner" },
  { key: "projectDate", label: "Project Date", kind: "date" },
  { key: "version", label: "Version" },
  { key: "overview", label: "Overview" },
];

const sheets: SheetSpec[] = [
  { name: "Business Case", moduleKey: "business-case", includeProject: true, sectionFields: [
    { key: "problem", label: "Problem or Opportunity" }, { key: "options", label: "Options Considered" },
    { key: "benefits", label: "Expected Benefits" }, { key: "costs", label: "Estimated Costs" },
    { key: "recommendation", label: "Recommendation" },
  ] },
  { name: "Project Charter", moduleKey: "charter", sectionFields: [
    { key: "purpose", label: "Purpose and Authority" }, { key: "successCriteria", label: "Success Criteria" }, { key: "approvals", label: "Approvals" },
  ], rowFields: [
    { key: "objectiveId", label: "Objective ID" }, { key: "objective", label: "SMART Objective" },
    { key: "kpi", label: "Measure / KPI" }, { key: "target", label: "Target" }, { key: "owner", label: "Owner" },
  ] },
  { name: "Scope Statement", moduleKey: "scope", sectionFields: [
    { key: "inScope", label: "In Scope" }, { key: "outOfScope", label: "Out of Scope" },
    { key: "deliverables", label: "Deliverables" }, { key: "exclusions", label: "Exclusions" },
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
    { key: "type", label: "Type" }, { key: "itemId", label: "ID" }, { key: "description", label: "Description" },
    { key: "owner", label: "Owner" }, { key: "date", label: "Review Date", kind: "date" }, { key: "impact", label: "Impact" }, { key: "status", label: "Status" },
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
    { key: "optimistic", label: "Optimistic (O)" }, { key: "most_likely", label: "Most Likely (M)" }, { key: "pessimistic", label: "Pessimistic (P)" },
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
    if (!("most_likely" in output) && "mostLikely" in output) output.most_likely = output.mostLikely ?? "";
    return output;
  } catch {
    return {};
  }
}

function normalizeHeader(value: unknown): string {
  return String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}

function cellText(value: unknown, field?: WorkbookField): string {
  if (value === null || value === undefined) return "";
  if (field?.kind === "date") {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      const y = value.getUTCFullYear();
      const m = String(value.getUTCMonth() + 1).padStart(2, "0");
      const d = String(value.getUTCDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return `${String(parsed.y).padStart(4, "0")}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
    const raw = String(value).trim();
    const iso = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) return `${iso[1]}-${String(iso[2]).padStart(2, "0")}-${String(iso[3]).padStart(2, "0")}`;
    const us = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (us) return `${us[3]}-${String(us[1]).padStart(2, "0")}-${String(us[2]).padStart(2, "0")}`;
    return raw;
  }
  if (field?.kind === "checkbox") {
    const normalized = String(value).trim().toLowerCase();
    return ["true", "yes", "y", "1", "x", "checked", "complete", "done"].includes(normalized) ? "true" : "false";
  }
  return String(value).trim();
}

function displayCell(value: string, field: WorkbookField): string | number {
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

function fileStem(value: string): string {
  const stem = value.trim().replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/^-+|-+$/g, "");
  return stem || "project";
}

export type WorkbookDownload = {
  fileName: string;
  blob: Blob;
};

function workbookDownload(workbook: XLSX.WorkBook, fileName: string): WorkbookDownload {
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
  return {
    fileName,
    blob: new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
  };
}

export function createProjectWorkbookDownload(data: WorkbookExportData): WorkbookDownload {
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
    const base = { ...(spec.includeProject ? data.project : {}), ...section };
    let records: Array<Record<string, string>>;
    if (spec.rowFields) records = moduleRows.length ? moduleRows.map((row, index) => ({ ...(index === 0 ? base : {}), ...row })) : Object.values(base).some(Boolean) ? [base] : [];
    else records = Object.values(base).some(Boolean) ? [base] : [];
    XLSX.utils.book_append_sheet(workbook, makeSheet(headers, records), spec.name);
  }
  const fileName = `${fileStem(data.project.projectCode || data.project.name)}-project-workbook.xlsx`;
  return workbookDownload(workbook, fileName);
}

export function createBlankWorkbookDownload(): WorkbookDownload {
  const workbook = XLSX.utils.book_new();
  for (const spec of sheets) {
    let headers = [...(spec.includeProject ? projectFields : []), ...(spec.sectionFields ?? []), ...(spec.rowFields ?? [])];
    if (spec.moduleKey === "raci") headers = [...headers, ...Array.from({ length: 8 }, (_, index) => ({ key: `role${index}`, label: `Role ${index + 1} Assignment` }))];
    XLSX.utils.book_append_sheet(workbook, makeSheet(headers, []), spec.name);
  }
  return workbookDownload(workbook, "project-workbook-template.xlsx");
}

function sheetRows(sheet: XLSX.WorkSheet): unknown[][] {
  return XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true, blankrows: false });
}

function readFields(row: unknown[], headerIndex: Map<string, number>, fields: WorkbookField[]): Record<string, string> {
  const output: Record<string, string> = {};
  for (const field of fields) {
    const index = headerIndex.get(normalizeHeader(field.label));
    output[field.key] = index === undefined ? "" : cellText(row[index], field);
  }
  return output;
}

function hasValue(record: Record<string, string>): boolean {
  return Object.values(record).some((value) => value.trim() !== "");
}

export async function parseProjectWorkbook(file: File): Promise<WorkbookImportPreview> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const notes: string[] = [];
  const counts: Array<{ sheet: string; count: number }> = [];
  const sections: SectionRecord[] = [];
  const rows: Array<Omit<RowRecord, "id">> = [];
  const project: ProjectRecord = { projectCode: "", name: "", sponsor: "", projectManager: "", businessOwner: "", projectDate: "", version: "1.0", overview: "" };
  const recognizedNames = new Set(sheets.map((spec) => normalizeHeader(spec.name)));
  for (const actualName of workbook.SheetNames) {
    if (!recognizedNames.has(normalizeHeader(actualName))) notes.push(`Skipped unknown sheet “${actualName}”.`);
  }

  for (const spec of sheets) {
    const actualName = workbook.SheetNames.find((name) => normalizeHeader(name) === normalizeHeader(spec.name));
    if (!actualName) {
      counts.push({ sheet: spec.name, count: 0 });
      notes.push(`${spec.name} was not found; this optional sheet will be left empty.`);
      continue;
    }
    const worksheet = workbook.Sheets[actualName];
    if (!worksheet) {
      counts.push({ sheet: spec.name, count: 0 });
      notes.push(`${spec.name} could not be read and was skipped.`);
      continue;
    }
    const matrix = sheetRows(worksheet);
    const header = matrix[0] ?? [];
    const headerIndex = new Map<string, number>();
    header.forEach((value, index) => { const key = normalizeHeader(value); if (key && !headerIndex.has(key)) headerIndex.set(key, index); });
    const expected = [...(spec.includeProject ? projectFields : []), ...(spec.sectionFields ?? []), ...(spec.rowFields ?? [])];
    const missingHeaders = expected.filter((field) => !headerIndex.has(normalizeHeader(field.label))).map((field) => field.label);
    if (missingHeaders.length) notes.push(`${spec.name}: ignored missing column${missingHeaders.length === 1 ? "" : "s"} ${missingHeaders.join(", ")}.`);
    const dataRows = matrix.slice(1);
    let importedCount = 0;
    let skippedCount = 0;
    let sectionSaved = false;
    let rowOrder = 0;
    let raciRoles: string[] = [];

    for (const row of dataRows) {
      const projectData = spec.includeProject ? readFields(row, headerIndex, projectFields) : {};
      if (spec.includeProject && hasValue(projectData)) Object.assign(project, projectData);
      const sectionData = spec.sectionFields ? readFields(row, headerIndex, spec.sectionFields) : {};
      if (spec.moduleKey === "raci") {
        if (!raciRoles.length) {
          raciRoles = (sectionData.roles ?? "").split(",").map((role) => role.trim()).filter(Boolean).slice(0, 8);
          if (!raciRoles.length) {
            const fixed = new Set(expected.map((field) => normalizeHeader(field.label)));
            raciRoles = header.map((value) => String(value ?? "").trim()).filter((value) => value && !fixed.has(normalizeHeader(value)) && !/^role \d+ assignment$/i.test(value)).slice(0, 8);
            sectionData.roles = raciRoles.join(", ");
          }
        }
      }
      if (spec.sectionFields && !sectionSaved && hasValue(sectionData)) {
        sections.push({ moduleKey: spec.moduleKey, dataJson: JSON.stringify(sectionData) });
        sectionSaved = true;
      }
      let rowData = spec.rowFields ? readFields(row, headerIndex, spec.rowFields) : {};
      if (spec.moduleKey === "raci" && spec.rowFields) {
        rowData = readFields(row, headerIndex, spec.rowFields);
        raciRoles.forEach((role, index) => {
          const named = headerIndex.get(normalizeHeader(role));
          const generic = headerIndex.get(normalizeHeader(`Role ${index + 1} Assignment`));
          rowData[`role${index}`] = cellText(row[named ?? generic ?? -1]);
        });
      }
      const rowHasValue = spec.rowFields ? hasValue(rowData) : false;
      const anyKnownValue = hasValue(projectData) || hasValue(sectionData) || rowHasValue;
      if (spec.rowFields && rowHasValue) {
        rows.push({ moduleKey: spec.moduleKey, dataJson: JSON.stringify(rowData), sortOrder: rowOrder });
        rowOrder += 1;
      }
      if (anyKnownValue) importedCount += 1;
      else if (row.some((cell) => cellText(cell) !== "")) skippedCount += 1;
    }
    if (skippedCount) notes.push(`${spec.name}: skipped ${skippedCount} row${skippedCount === 1 ? "" : "s"} with no recognized data.`);
    if (spec.sectionFields && !sectionSaved && dataRows.length > 0) {
      const emptySection: Record<string, string> = {};
      for (const field of spec.sectionFields) emptySection[field.key] = "";
      if (spec.includeProject && (project.name || project.projectCode)) sections.push({ moduleKey: spec.moduleKey, dataJson: JSON.stringify(emptySection) });
    }
    counts.push({ sheet: spec.name, count: importedCount });
  }
  if (!project.version.trim()) project.version = "1.0";
  if (!project.name.trim()) notes.unshift("Project Name is required in the Business Case sheet before this workbook can be imported.");
  if (!project.projectCode.trim()) notes.unshift("Project ID is required in the Business Case sheet before this workbook can be imported.");
  return { fileName: file.name, project, sections, rows, counts, notes, canImport: Boolean(project.name.trim() && project.projectCode.trim()) };
}
