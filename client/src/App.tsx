import { useMemo, useRef, useState } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { SafeAreaTopScrim } from "@hatch/space-sdk/client";
import { api, type ApiResponse } from "./api";
import { parseProjectWorkbook, type WorkbookImportPreview } from "./excel";

type Project = ApiResponse<typeof api, "listProjects">["projects"][number];
type DataRow = ApiResponse<typeof api, "getProjectData">["rows"][number];
type FieldType = "text" | "textarea" | "date" | "number" | "select" | "range" | "checkbox";
type Field = { key: string; label: string; type?: FieldType; options?: string[]; placeholder?: string };
type ModuleDef = { key: string; label: string; phase: Phase; intro: string; fields?: Field[]; rowLabel?: string; rowFields?: Field[] };
type Phase = "INITIATE" | "PLAN" | "EXECUTE" | "MONITOR";

const projectFields: Field[] = [
  { key: "name", label: "Project Name" }, { key: "projectCode", label: "Project ID" },
  { key: "sponsor", label: "Sponsor" }, { key: "projectManager", label: "Project Manager" },
  { key: "businessOwner", label: "Business Owner" }, { key: "projectDate", label: "Date", type: "date" },
  { key: "version", label: "Version" }, { key: "overview", label: "Overview", type: "textarea" },
];

const statusOptions = ["Not Started", "In Progress", "Blocked", "Complete"];
const priorityOptions = ["Low", "Medium", "High"];
const moduleDefs: ModuleDef[] = [
  { key: "business-case", label: "Business Case", phase: "INITIATE", intro: "Frame the decision before committing resources.", fields: [
    { key: "problem", label: "Problem or Opportunity", type: "textarea" }, { key: "options", label: "Options Considered", type: "textarea" },
    { key: "benefits", label: "Expected Benefits", type: "textarea" }, { key: "costs", label: "Estimated Costs", type: "textarea" },
    { key: "recommendation", label: "Recommendation", type: "textarea" },
  ]},
  { key: "charter", label: "Project Charter", phase: "INITIATE", intro: "Authorize the work and define success.", fields: [
    { key: "purpose", label: "Purpose and Authority", type: "textarea" }, { key: "successCriteria", label: "Success Criteria", type: "textarea" },
    { key: "approvals", label: "Approvals", type: "textarea" },
  ], rowLabel: "Charter Objective", rowFields: [
    { key: "objectiveId", label: "Objective ID" }, { key: "objective", label: "SMART Objective", type: "textarea" },
    { key: "kpi", label: "Measure / KPI" }, { key: "target", label: "Target" }, { key: "owner", label: "Owner" },
  ]},
  { key: "stakeholders", label: "Stakeholder Register", phase: "INITIATE", intro: "Map influence, interest, and the engagement gap.", rowLabel: "Stakeholder", rowFields: [
    { key: "name", label: "Name" }, { key: "orgRole", label: "Organization / Role" },
    { key: "influence", label: "Influence", type: "select", options: ["Low", "Medium", "High"] },
    { key: "interest", label: "Interest", type: "select", options: ["Low", "Medium", "High"] },
    { key: "currentEngagement", label: "Current Engagement" }, { key: "desiredEngagement", label: "Desired Engagement" },
  ]},
  { key: "objectives", label: "Project Objectives", phase: "PLAN", intro: "Turn outcomes into owned, measurable targets.", rowLabel: "Objective", rowFields: [
    { key: "objective", label: "SMART Objective", type: "textarea" }, { key: "kpi", label: "KPI" }, { key: "target", label: "Target" }, { key: "owner", label: "Owner" },
  ]},
  { key: "requirements", label: "High-Level Requirements", phase: "PLAN", intro: "Capture what must be true for delivery to succeed.", rowLabel: "Requirement", rowFields: [
    { key: "requirementId", label: "Requirement ID" }, { key: "description", label: "Description", type: "textarea" },
    { key: "priority", label: "Priority", type: "select", options: priorityOptions }, { key: "owner", label: "Owner" },
    { key: "status", label: "Status", type: "select", options: statusOptions },
  ]},
  { key: "scope", label: "Scope Statement", phase: "PLAN", intro: "Draw the boundaries before the schedule hardens.", fields: [
    { key: "inScope", label: "In Scope", type: "textarea" }, { key: "outOfScope", label: "Out of Scope", type: "textarea" },
    { key: "deliverables", label: "Deliverables", type: "textarea" }, { key: "exclusions", label: "Exclusions", type: "textarea" },
  ]},
  { key: "wbs", label: "WBS Builder", phase: "PLAN", intro: "Break deliverables into estimable work packages.", rowLabel: "WBS Item", rowFields: [
    { key: "wbsId", label: "WBS ID" }, { key: "level", label: "Level", type: "number" }, { key: "deliverable", label: "Deliverable / Work Package" },
    { key: "requirementIds", label: "Linked Requirement IDs", placeholder: "REQ-01, REQ-02" },
    { key: "description", label: "Description", type: "textarea" }, { key: "owner", label: "Owner" }, { key: "estimateHours", label: "Estimate Hours", type: "number" },
    { key: "budgetedCost", label: "Budgeted Cost", type: "number", placeholder: "0.00" }, { key: "actualCost", label: "Actual Cost", type: "number", placeholder: "0.00" },
    { key: "start", label: "Start", type: "date" }, { key: "end", label: "End", type: "date" }, { key: "dependencies", label: "Dependencies" },
  ]},
  { key: "pert", label: "PERT Estimator", phase: "PLAN", intro: "Use three-point estimates to expose uncertainty.", rowLabel: "PERT Estimate", rowFields: [
    { key: "activityId", label: "Activity ID" }, { key: "activityName", label: "Activity Name" }, { key: "predecessors", label: "Predecessor(s)" },
    { key: "optimistic", label: "Optimistic (O)", type: "number" }, { key: "mostLikely", label: "Most Likely (M)", type: "number" },
    { key: "pessimistic", label: "Pessimistic (P)", type: "number" }, { key: "owner", label: "Owner" }, { key: "notes", label: "Notes", type: "textarea" },
  ]},
  { key: "timeline", label: "Project Timeline", phase: "PLAN", intro: "Build the dependency-aware baseline.", rowLabel: "Timeline Item", rowFields: [
    { key: "wbsId", label: "WBS ID" }, { key: "taskName", label: "Task Name" }, { key: "owner", label: "Owner" },
    { key: "startDate", label: "Start Date", type: "date" }, { key: "endDate", label: "End Date", type: "date" }, { key: "duration", label: "Duration (days)", type: "number" },
    { key: "percentComplete", label: "% Complete", type: "range" }, { key: "status", label: "Status", type: "select", options: statusOptions },
    { key: "predecessors", label: "Predecessor(s)" }, { key: "dependencyType", label: "Dependency Type", type: "select", options: ["FS", "SS", "FF", "SF"] },
    { key: "milestone", label: "Milestone", type: "select", options: ["No", "Yes"] }, { key: "notes", label: "Notes", type: "textarea" },
  ]},
  { key: "raci", label: "RACI Matrix", phase: "PLAN", intro: "Assign accountability before work starts.", fields: [{ key: "roles", label: "Roles (comma separated)", placeholder: "Project Manager, Superintendent, Owner" }], rowLabel: "Activity", rowFields: [{ key: "activity", label: "Activity" }] },
  { key: "management-plan", label: "Project Mgmt Plan", phase: "PLAN", intro: "Set the operating system for delivery.", fields: [
    { key: "approach", label: "Delivery Approach", type: "textarea" }, { key: "governance", label: "Governance", type: "textarea" }, { key: "cadence", label: "Meeting and Reporting Cadence", type: "textarea" },
  ]},
  { key: "ac-log", label: "Assumptions & Constraints", phase: "PLAN", intro: "Keep the conditions shaping the plan visible.", rowLabel: "Log Entry", rowFields: [
    { key: "type", label: "Type", type: "select", options: ["Assumption", "Constraint"] }, { key: "itemId", label: "ID" },
    { key: "description", label: "Description", type: "textarea" }, { key: "owner", label: "Owner" }, { key: "date", label: "Review Date", type: "date" },
    { key: "impact", label: "Impact", type: "select", options: ["Low", "Medium", "High"] }, { key: "status", label: "Status", type: "select", options: ["Open", "Closed"] },
  ]},
  { key: "kickoff", label: "Kickoff Agenda", phase: "PLAN", intro: "Close planning with a disciplined handoff.", rowLabel: "Agenda Item", rowFields: [
    { key: "complete", label: "Complete", type: "checkbox" }, { key: "item", label: "Agenda Item" }, { key: "owner", label: "Owner" }, { key: "notes", label: "Decisions / Notes", type: "textarea" },
  ]},
  { key: "task-tracker", label: "Task Tracker", phase: "EXECUTE", intro: "Run the work from clear ownership and visible status.", rowLabel: "Task", rowFields: [
    { key: "taskId", label: "Task ID" }, { key: "task", label: "Task" }, { key: "owner", label: "Owner" },
    { key: "startDate", label: "Start Date", type: "date" }, { key: "dueDate", label: "Due Date", type: "date" },
    { key: "priority", label: "Priority", type: "select", options: priorityOptions }, { key: "status", label: "Status", type: "select", options: statusOptions },
    { key: "percentComplete", label: "% Complete", type: "range" }, { key: "blocker", label: "Blocker / Risk" }, { key: "nextStep", label: "Next Step" },
    { key: "dependency", label: "Dependency" }, { key: "notes", label: "Notes", type: "textarea" },
  ]},
  { key: "task-manager", label: "Task Manager", phase: "EXECUTE", intro: "Track effort, readiness, dependencies, and done criteria.", rowLabel: "Detailed Task", rowFields: [
    { key: "taskId", label: "Task ID" }, { key: "taskName", label: "Task Name" }, { key: "owner", label: "Owner" }, { key: "workstream", label: "Workstream" },
    { key: "start", label: "Start", type: "date" }, { key: "due", label: "Due", type: "date" }, { key: "priority", label: "Priority", type: "select", options: priorityOptions },
    { key: "status", label: "Status", type: "select", options: statusOptions }, { key: "effort", label: "Effort (hrs)", type: "number" }, { key: "actual", label: "Actual (hrs)", type: "number" },
    { key: "predecessors", label: "Predecessor(s)" }, { key: "readyCriteria", label: "Ready Criteria", type: "textarea" }, { key: "doneCriteria", label: "Done Criteria", type: "textarea" },
  ]},
  { key: "todo", label: "To-Do List", phase: "EXECUTE", intro: "Keep the immediate commitments short and explicit.", rowLabel: "To-Do", rowFields: [
    { key: "done", label: "Done", type: "checkbox" }, { key: "item", label: "Item" }, { key: "priority", label: "Priority", type: "select", options: priorityOptions },
    { key: "dueDate", label: "Due Date", type: "date" }, { key: "estimatedMinutes", label: "Estimated Minutes", type: "number" }, { key: "category", label: "Category" }, { key: "notes", label: "Notes", type: "textarea" },
  ]},
  { key: "work-schedule", label: "Work Schedule", phase: "EXECUTE", intro: "Balance weekly planned hours against actual effort.", fields: [{ key: "weekStarting", label: "Week Starting", type: "date" }], rowLabel: "Resource", rowFields: [
    { key: "resource", label: "Resource" }, { key: "role", label: "Role" }, ...["mon","tue","wed","thu","fri","sat","sun"].map((key) => ({ key, label: `${key[0]?.toUpperCase() ?? ""}${key.slice(1)} (hrs)`, type: "number" as FieldType })),
    { key: "actualTotal", label: "Actual Total", type: "number" }, { key: "notes", label: "Notes", type: "textarea" },
  ]},
  { key: "calendar", label: "Project Calendar", phase: "EXECUTE", intro: "Keep meetings, milestones, and field events in one list.", rowLabel: "Calendar Event", rowFields: [
    { key: "date", label: "Date", type: "date" }, { key: "startTime", label: "Start Time" }, { key: "endTime", label: "End Time" },
    { key: "title", label: "Event / Meeting" }, { key: "category", label: "Category" }, { key: "owner", label: "Organizer / Owner" },
    { key: "attendees", label: "Attendees" }, { key: "location", label: "Location / Link" }, { key: "description", label: "Description", type: "textarea" },
  ]},
  { key: "automation", label: "Attention & Reports", phase: "MONITOR", intro: "Validate accountability, scan schedule health, review budget variance, and prepare the weekly status draft from saved project data." },
  { key: "dashboard", label: "Project Health", phase: "MONITOR", intro: "Read progress directly from the task tracker." },
];

const phases: Phase[] = ["INITIATE", "PLAN", "EXECUTE", "MONITOR"];
const phaseNames: Record<Phase, string> = { INITIATE: "Initiate", PLAN: "Plan", EXECUTE: "Execute", MONITOR: "Monitor" };

function parseJson(value: string | undefined): Record<string, string> {
  try { return value ? JSON.parse(value) as Record<string, string> : {}; } catch { return {}; }
}

function FieldControl({ field, value, onChange }: { field: Field; value: string; onChange: (value: string) => void }) {
  const id = `${field.key}-${Math.random().toString(36).slice(2)}`;
  if (field.type === "textarea") return <label className="field"><span>{field.label}</span><textarea aria-label={field.label} value={value} placeholder={field.placeholder} onChange={(e) => onChange(e.target.value)} /></label>;
  if (field.type === "select") return <label className="field"><span>{field.label}</span><select aria-label={field.label} value={value} onChange={(e) => onChange(e.target.value)}><option value="">Select</option>{field.options?.map((o) => <option key={o}>{o}</option>)}</select></label>;
  if (field.type === "checkbox") return <label className="check-field" htmlFor={id}><input id={id} aria-label={field.label} type="checkbox" checked={value === "true"} onChange={(e) => onChange(String(e.target.checked))} /><span>{field.label}</span></label>;
  if (field.type === "range") return <label className="field range-field"><span>{field.label} <b>{value || "0"}%</b></span><input aria-label={field.label} type="range" min="0" max="100" step="5" value={value || "0"} onChange={(e) => onChange(e.target.value)} /></label>;
  return <label className="field"><span>{field.label}</span><input aria-label={field.label} type={field.type ?? "text"} value={value} placeholder={field.placeholder} step={field.type === "number" ? "0.01" : undefined} onChange={(e) => onChange(e.target.value)} /></label>;
}

function CreateProject({ onCreated, compact = false }: { onCreated: (id: number) => void; compact?: boolean }) {
  const qc = useQueryClient();
  const [form, setForm] = useState<Record<string, string>>({ version: "1.0" });
  const [message, setMessage] = useState("");
  const create = useMutation({ mutationFn: () => api.createProject({
    projectCode: form.projectCode ?? "", name: form.name ?? "", sponsor: form.sponsor ?? "", projectManager: form.projectManager ?? "",
    businessOwner: form.businessOwner ?? "", projectDate: form.projectDate ?? "", version: form.version ?? "1.0", overview: form.overview ?? "",
  }), onSuccess: async (result) => { setMessage(result.message); if (result.ok && result.id) { await qc.invalidateQueries(); onCreated(result.id); } } });
  return <form className={compact ? "create-compact" : "empty-state"} onSubmit={(e) => { e.preventDefault(); setMessage(""); create.mutate(); }}>
    {!compact && <><p className="eyebrow">First job</p><h2>Create your first project</h2><p>Start with the project header. Every phase and worksheet stays isolated to this project.</p></>}
    <div className="form-grid">{projectFields.map((f) => <FieldControl key={f.key} field={f} value={form[f.key] ?? ""} onChange={(v) => setForm((old) => ({ ...old, [f.key]: v }))} />)}</div>
    <div className="action-row"><button className="primary" disabled={create.isPending}>{create.isPending ? "Creating…" : "Create project"}</button>{message && <p className={message.includes("created") ? "success" : "error"}>{message}</p>}</div>
  </form>;
}

function ProjectHeaderForm({ project }: { project: Project }) {
  const qc = useQueryClient();
  const [form, setForm] = useState<Record<string, string>>({ projectCode: project.projectCode, name: project.name, sponsor: project.sponsor, projectManager: project.projectManager, businessOwner: project.businessOwner, projectDate: project.projectDate, version: project.version, overview: project.overview });
  const [message, setMessage] = useState("");
  const save = useMutation({ mutationFn: () => api.updateProject({ id: project.id, projectCode: form.projectCode ?? "", name: form.name ?? "", sponsor: form.sponsor ?? "", projectManager: form.projectManager ?? "", businessOwner: form.businessOwner ?? "", projectDate: form.projectDate ?? "", version: form.version ?? "", overview: form.overview ?? "" }), onSuccess: async (r) => { setMessage(r.message); await qc.invalidateQueries(); } });
  return <details className="project-header"><summary>Project header <span>{project.projectCode}</span></summary><form onSubmit={(e) => { e.preventDefault(); save.mutate(); }}><div className="form-grid">{projectFields.map((f) => <FieldControl key={f.key} field={f} value={form[f.key] ?? ""} onChange={(v) => setForm((old) => ({ ...old, [f.key]: v }))} />)}</div><div className="action-row"><button className="primary" disabled={save.isPending}>{save.isPending ? "Saving…" : "Save header"}</button>{message && <p className="success">{message}</p>}</div></form></details>;
}

function SectionForm({ projectId, module, dataJson, updatedAt }: { projectId: number; module: ModuleDef; dataJson?: string; updatedAt?: string }) {
  const qc = useQueryClient();
  const [form, setForm] = useState<Record<string, string>>(() => parseJson(dataJson));
  const [message, setMessage] = useState("");
  const save = useMutation({ mutationFn: () => api.saveSection({ projectId, moduleKey: module.key, dataJson: JSON.stringify(form) }), onSuccess: async (r) => { setMessage(r.message); await qc.invalidateQueries(); } });
  if (!module.fields?.length) return null;
  return <form className="section-form" data-version={updatedAt} onSubmit={(e) => { e.preventDefault(); save.mutate(); }}><div className="form-grid">{module.fields.map((f) => <FieldControl key={f.key} field={f} value={form[f.key] ?? ""} onChange={(v) => setForm((old) => ({ ...old, [f.key]: v }))} />)}</div><div className="action-row"><button className="primary" disabled={save.isPending}>{save.isPending ? "Saving…" : "Save section"}</button>{message && <p className={save.data?.ok ? "success" : "error"}>{message}</p>}</div></form>;
}

function RowEditor({ row, projectId, module, raciRoles, allRows }: { row: DataRow; projectId: number; module: ModuleDef; raciRoles: string[]; allRows: DataRow[] }) {
  const qc = useQueryClient();
  const [open, setOpen] = useState(false);
  const [spawnOpen, setSpawnOpen] = useState(false);
  const [form, setForm] = useState<Record<string, string>>(() => {
    const initial = parseJson(row.dataJson);
    if (module.key === "ac-log" && !initial.status) initial.status = "Open";
    return initial;
  });
  const [spawnOwner, setSpawnOwner] = useState(() => parseJson(row.dataJson).owner ?? "");
  const [spawnDueDate, setSpawnDueDate] = useState("");
  const [message, setMessage] = useState("");
  const fields = module.key === "raci" ? [...(module.rowFields ?? []), ...raciRoles.map((role, i) => ({ key: `role${i}`, label: role, type: "select" as FieldType, options: ["R", "A", "C", "I"] }))] : (module.rowFields ?? []);
  const title = form.task || form.taskName || form.deliverable || form.activityName || form.activity || form.name || form.item || form.title || form.description || `${module.rowLabel} ${row.id}`;
  const save = useMutation({ mutationFn: () => api.updateRow({ id: row.id, projectId, dataJson: JSON.stringify(form) }), onSuccess: async (r) => { setMessage(r.message); await qc.invalidateQueries(); } });
  const remove = useMutation({ mutationFn: () => api.deleteRow({ id: row.id, projectId }), onSuccess: async () => qc.invalidateQueries() });
  const promote = useMutation({ mutationFn: () => api.promoteWbs({ projectId, rowId: row.id }), onSuccess: async (r) => { setMessage(r.message); await qc.invalidateQueries(); } });
  const sendPert = useMutation({ mutationFn: () => api.sendPertToTimeline({ projectId, rowId: row.id }), onSuccess: async (r) => { setMessage(r.message); await qc.invalidateQueries(); } });
  const spawnKickoff = useMutation({ mutationFn: () => api.spawnKickoffTask({ projectId, rowId: row.id, owner: spawnOwner, dueDate: spawnDueDate }), onSuccess: async (r) => { setMessage(r.message); if (r.ok) setSpawnOpen(false); await qc.invalidateQueries(); } });
  const linkedTask = module.key === "kickoff" ? allRows.find((candidate) => candidate.moduleKey === "task-tracker" && candidate.sourceRowId === row.id) : undefined;
  const linkedTaskData = linkedTask ? parseJson(linkedTask.dataJson) : undefined;
  const o = Number(form.optimistic || 0), m = Number(form.mostLikely || 0), p = Number(form.pessimistic || 0);
  const expected = (o + 4 * m + p) / 6, sd = (p - o) / 6;
  const planned = ["mon","tue","wed","thu","fri","sat","sun"].reduce((sum, key) => sum + Number(form[key] || 0), 0);
  const hasBudget = form.budgetedCost !== "" && form.budgetedCost !== undefined;
  const hasActual = form.actualCost !== "" && form.actualCost !== undefined;
  const budget = Number(form.budgetedCost || 0), actual = Number(form.actualCost || 0), variance = actual - budget;
  const money = (value: number) => new Intl.NumberFormat(undefined, { style: "currency", currency: "USD" }).format(value);
  return <article className="data-row"><button className="row-summary" type="button" onClick={() => setOpen(!open)} aria-expanded={open}><span>{title}</span><b>{open ? "Close" : "Edit"}</b></button>
    {module.key === "pert" && <div className="calc-strip"><span>TE <b>{expected.toFixed(2)} days</b></span><span>σ <b>{sd.toFixed(2)}</b></span><span>Variance <b>{(sd * sd).toFixed(2)}</b></span></div>}
    {module.key === "work-schedule" && <div className="calc-strip"><span>Planned <b>{planned.toFixed(1)} hrs</b></span><span>Actual <b>{Number(form.actualTotal || 0).toFixed(1)} hrs</b></span></div>}
    {module.key === "wbs" && (hasBudget || hasActual) && <div className="calc-strip budget-strip"><span>Budget <b>{money(budget)}</b></span><span>Actual <b>{money(actual)}</b></span><span className={variance > 0 ? "over" : variance < 0 ? "under" : "even"}>Variance <b>{money(variance)}</b> · {variance > 0 ? "Over" : variance < 0 ? "Under" : "On budget"}</span></div>}
    {module.key === "kickoff" && linkedTask && <div className="linked-strip" aria-label="Linked tracked task"><b>Spawned</b><span>{linkedTaskData?.task || `Task ${linkedTask.id}`}{linkedTaskData?.owner ? ` · ${linkedTaskData.owner}` : ""}{linkedTaskData?.dueDate ? ` · due ${linkedTaskData.dueDate}` : ""}</span></div>}
    {open && <form className="row-form" onSubmit={(e) => { e.preventDefault(); save.mutate(); }}><div className="form-grid">{fields.map((f) => <FieldControl key={f.key} field={f} value={form[f.key] ?? ""} onChange={(v) => setForm((old) => ({ ...old, [f.key]: v }))} />)}</div><div className="action-row"><button className="primary" disabled={save.isPending}>{save.isPending ? "Saving…" : "Save row"}</button>{module.key === "wbs" && <button type="button" className="secondary" disabled={promote.isPending} onClick={() => promote.mutate()}>{promote.isPending ? "Promoting…" : "Promote to task tracker"}</button>}{module.key === "pert" && <button type="button" className="secondary" disabled={sendPert.isPending} onClick={() => sendPert.mutate()}>{sendPert.isPending ? "Sending…" : "Add TE to timeline"}</button>}{module.key === "kickoff" && !linkedTask && <button type="button" className="secondary" onClick={() => { setSpawnOwner(form.owner ?? ""); setSpawnOpen(!spawnOpen); }}>{spawnOpen ? "Cancel task setup" : "Spawn tracked task"}</button>}<button type="button" className="danger" disabled={remove.isPending} onClick={() => remove.mutate()}>{remove.isPending ? "Removing…" : "Remove"}</button>{message && <p className={(save.data?.ok ?? promote.data?.ok ?? sendPert.data?.ok ?? spawnKickoff.data?.ok) ? "success" : "error"}>{message}</p>}</div></form>}
    {open && module.key === "kickoff" && spawnOpen && !linkedTask && <form className="spawn-form" onSubmit={(event) => { event.preventDefault(); setMessage(""); spawnKickoff.mutate(); }}><div><strong>Create tracked task</strong><span>The agenda item text will become the task name.</span></div><div className="form-grid"><FieldControl field={{ key: "spawnOwner", label: "Task Owner" }} value={spawnOwner} onChange={setSpawnOwner} /><FieldControl field={{ key: "spawnDueDate", label: "Task Due Date", type: "date" }} value={spawnDueDate} onChange={setSpawnDueDate} /></div><div className="action-row"><button className="primary" disabled={spawnKickoff.isPending}>{spawnKickoff.isPending ? "Creating task…" : "Create and link task"}</button>{spawnKickoff.isError && <p className="error">The task could not be created. Try again.</p>}</div></form>}
  </article>;
}

function AddRow({ projectId, module, raciRoles }: { projectId: number; module: ModuleDef; raciRoles: string[] }) {
  const qc = useQueryClient();
  const [open, setOpen] = useState(false);
  const [form, setForm] = useState<Record<string, string>>(module.key === "ac-log" ? { status: "Open" } : {});
  const [message, setMessage] = useState("");
  const fields = module.key === "raci" ? [...(module.rowFields ?? []), ...raciRoles.map((role, i) => ({ key: `role${i}`, label: role, type: "select" as FieldType, options: ["R", "A", "C", "I"] }))] : (module.rowFields ?? []);
  const add = useMutation({ mutationFn: () => api.addRow({ projectId, moduleKey: module.key, dataJson: JSON.stringify(form) }), onSuccess: async (r) => { setMessage(r.message); if (r.ok) { setForm(module.key === "ac-log" ? { status: "Open" } : {}); setOpen(false); } await qc.invalidateQueries(); } });
  if (!module.rowFields) return null;
  return <div className="add-row"><button type="button" className="secondary add-button" onClick={() => setOpen(!open)}>{open ? "Cancel" : `Add ${module.rowLabel?.toLowerCase()}`}</button>{open && <form className="row-form" onSubmit={(e) => { e.preventDefault(); add.mutate(); }}><div className="form-grid">{fields.map((f) => <FieldControl key={f.key} field={f} value={form[f.key] ?? ""} onChange={(v) => setForm((old) => ({ ...old, [f.key]: v }))} />)}</div><div className="action-row"><button className="primary" disabled={add.isPending}>{add.isPending ? "Adding…" : `Add ${module.rowLabel}`}</button>{message && <p className={add.data?.ok ? "success" : "error"}>{message}</p>}</div></form>}</div>;
}

function localDateValue(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function ExcelTools({ selected, onImported }: { selected: Project; onImported: (id: number) => void }) {
  const qc = useQueryClient();
  const [parsing, setParsing] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  const [preview, setPreview] = useState<WorkbookImportPreview | null>(null);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const importWorkbook = useMutation({
    mutationFn: (value: WorkbookImportPreview) => api.importProjectWorkbook({ project: value.project, sections: value.sections, rows: value.rows }),
    onSuccess: async (result) => {
      if (!result.ok || !result.id) { setError(result.message); return; }
      setMessage(`${result.message} ${result.sectionsCreated} sections and ${result.rowsCreated} rows imported.`);
      setPreview(null);
      await qc.invalidateQueries();
      onImported(result.id);
    },
    onError: () => setError("The project could not be created from this workbook. Your saved projects were not changed."),
  });
  const projectExport = useQuery({
    queryKey: ["workbook-export", selected.id, selected.updatedAt],
    queryFn: () => api.prepareWorkbookExport({ projectId: selected.id, template: false }),
    staleTime: Infinity,
  });
  const blankExport = useQuery({
    queryKey: ["workbook-export", "blank-template"],
    queryFn: () => api.prepareWorkbookExport({ template: true }),
    staleTime: Infinity,
  });
  const projectFile = projectExport.data?.ok && projectExport.data.url && projectExport.data.fileName ? projectExport.data : null;
  const blankFile = blankExport.data?.ok && blankExport.data.url && blankExport.data.fileName ? blankExport.data : null;
  const exportError = projectExport.isError || blankExport.isError
    ? "The Excel file could not be prepared. Reopen the artifact and try again."
    : projectExport.data && !projectExport.data.ok
      ? projectExport.data.message
      : blankExport.data && !blankExport.data.ok
        ? blankExport.data.message
        : "";
  const selectFile = async (file: File | undefined) => {
    if (!file) return;
    setParsing(true); setMessage(""); setError(""); setPreview(null);
    if (!file.name.toLowerCase().endsWith(".xlsx")) { setError("Choose an .xlsx workbook."); setParsing(false); return; }
    await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
    try { setPreview(await parseProjectWorkbook(file)); }
    catch { setError("This .xlsx file could not be read. Check that it is a valid Excel workbook and try again."); }
    finally { setParsing(false); }
  };
  return <section className="excel-tools" aria-labelledby="excel-tools-title">
    <div className="excel-tools-copy"><p>Workbook exchange</p><h2 id="excel-tools-title">Excel import & export</h2><span>Round-trip the active project or start from a clean 16-sheet template.</span></div>
    <div className="excel-actions">
      {projectFile ? <a className="primary download-button" href={projectFile.url} target="_blank" rel="noopener noreferrer">Export to Excel</a> : <button type="button" className="primary" disabled>{projectExport.isError || projectExport.data?.ok === false ? "Export unavailable" : "Preparing export…"}</button>}
      {blankFile ? <a className="secondary download-button" href={blankFile.url} target="_blank" rel="noopener noreferrer">Export blank template</a> : <button type="button" className="secondary" disabled>{blankExport.isError || blankExport.data?.ok === false ? "Template unavailable" : "Preparing template…"}</button>}
      <button type="button" className="secondary" disabled={parsing || importWorkbook.isPending} onClick={() => fileInput.current?.click()}>{parsing ? "Reading workbook…" : "Import from Excel"}</button>
      <input ref={fileInput} className="file-input-hidden" tabIndex={-1} aria-hidden="true" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" disabled={importWorkbook.isPending} onChange={(event) => { const file = event.target.files?.[0]; event.target.value = ""; void selectFile(file); }} />
    </div>
    <p className="excel-delivery-note">Exports open as real Excel files in Muse. On Mac, save the opened file to Downloads. On iPhone or iPad, use Share, then Save to Files.</p>
    {(message || error || exportError) && <p className={(error || exportError) ? "error excel-message" : "success excel-message"} aria-live="polite">{error || exportError || message}</p>}
    {preview && <div className="import-overlay" role="presentation"><section className="import-preview" role="dialog" aria-modal="true" aria-labelledby="import-preview-title">
      <div className="import-preview-head"><div><p>Import preview</p><h3 id="import-preview-title">{preview.project.name || "Project name not detected"}</h3><span>{preview.fileName}{preview.project.projectCode ? ` · ${preview.project.projectCode}` : ""}</span></div><button type="button" className="preview-close" aria-label="Close import preview" disabled={importWorkbook.isPending} onClick={() => setPreview(null)}>Close</button></div>
      <div className="import-summary"><div><span>Sections</span><b>{preview.sections.length}</b></div><div><span>Rows</span><b>{preview.rows.length}</b></div><div><span>Sheets found</span><b>{preview.counts.filter((item) => item.count > 0).length}</b></div></div>
      <div className="sheet-counts" aria-label="Rows detected by worksheet">{preview.counts.map((item) => <div key={item.sheet}><span>{item.sheet}</span><b>{item.count}</b></div>)}</div>
      {preview.notes.length > 0 && <div className="import-notes"><strong>Import notes</strong><ul>{preview.notes.map((note, index) => <li key={`${note}-${index}`}>{note}</li>)}</ul></div>}
      <div className="import-confirm"><p>A new project will be created. Existing projects will not be changed.</p><button type="button" className="primary" disabled={!preview.canImport || importWorkbook.isPending} onClick={() => { setError(""); importWorkbook.mutate(preview); }}>{importWorkbook.isPending ? "Creating project…" : "Confirm and create project"}</button>{!preview.canImport && <span>Add Project Name and Project ID in the Business Case sheet, then upload the workbook again.</span>}{error && <span className="error">{error}</span>}</div>
    </section></div>}
  </section>;
}

function AutomationCenter({ projectId, rows, onNavigate }: { projectId: number; rows: DataRow[]; onNavigate: (phase: Phase, moduleKey: string) => void }) {
  const [reportDate, setReportDate] = useState(localDateValue);
  const [draft, setDraft] = useState("");
  const [copyState, setCopyState] = useState<"idle" | "copying" | "copied" | "error">("idle");
  const trace = useMutation({ mutationFn: () => api.runTraceabilityCheck({ projectId }) });
  const raci = useMutation({ mutationFn: () => api.runRaciValidation({ projectId }) });
  const schedule = useMutation({ mutationFn: () => api.runScheduleHealthScan({ projectId, scanDate: localDateValue() }) });
  const compile = useMutation({ mutationFn: () => api.compileStatusReport({ projectId, reportDate }), onSuccess: (result) => { setDraft(result.draft); setCopyState("idle"); } });
  const wbsCosts = rows.filter((row) => row.moduleKey === "wbs").map((row) => ({ row, data: parseJson(row.dataJson) })).filter(({ data }) => data.budgetedCost !== "" && data.budgetedCost !== undefined || data.actualCost !== "" && data.actualCost !== undefined);
  const budgetSummary = wbsCosts.map(({ row, data }) => {
    const budget = Number(data.budgetedCost || 0), actual = Number(data.actualCost || 0);
    return { rowId: row.id, name: data.deliverable?.trim() || data.wbsId?.trim() || `WBS item ${row.id}`, budget, actual, variance: actual - budget };
  });
  const totalBudget = budgetSummary.reduce((sum, item) => sum + item.budget, 0);
  const totalActual = budgetSummary.reduce((sum, item) => sum + item.actual, 0);
  const totalVariance = totalActual - totalBudget;
  const money = (value: number) => new Intl.NumberFormat(undefined, { style: "currency", currency: "USD" }).format(value);
  const copyDraft = async () => {
    if (!draft) return;
    setCopyState("copying");
    try { await navigator.clipboard.writeText(draft); setCopyState("copied"); }
    catch { setCopyState("error"); }
  };
  const checkedAt = trace.data?.checkedAt ? new Date(trace.data.checkedAt).toLocaleString() : "";
  return <div className="automation-grid">
    <div className="monitor-tools">
      <section className="automation-tool attention-tool" aria-labelledby="traceability-title">
        <div className="tool-heading"><div><p className="tool-kicker">Traceability</p><h3 id="traceability-title">Attention needed</h3></div>{trace.data && <span className={trace.data.gaps.length ? "gap-count" : "clear-count"}>{trace.data.gaps.length}</span>}</div>
        <p className="tool-copy">Check requirement coverage, WBS promotion, and task ownership against the latest saved records.</p>
        <button type="button" className="primary wide-action" disabled={trace.isPending} onClick={() => trace.mutate()}>{trace.isPending ? "Checking project flow…" : trace.data ? "Run check again" : "Run traceability check"}</button>
        {trace.isError && <p className="error result-message">The check could not run. Try again.</p>}
        {trace.data && <div className="trace-results" aria-live="polite">
          <div className="check-summary"><strong>{trace.data.message}</strong><span>{trace.data.requirementsChecked} requirements · {trace.data.wbsChecked} WBS items · {trace.data.tasksChecked} tasks</span><small>Checked {checkedAt}</small></div>
          {!trace.data.gaps.length ? <div className="clear-state"><b>Coverage is clear</b><span>Every saved requirement is linked, every WBS item is promoted, and every task has an owner.</span></div> : <div className="gap-list">{trace.data.gaps.map((gap) => <button key={`${gap.type}-${gap.rowId}`} type="button" className="gap-item" onClick={() => onNavigate(gap.moduleKey === "task-tracker" ? "EXECUTE" : "PLAN", gap.moduleKey)}><span className={`gap-type ${gap.type}`}>{gap.type === "requirement" ? "Requirement" : gap.type === "wbs" ? "WBS" : "Task"}</span><strong>{gap.label}</strong><small>{gap.detail}</small><i>Open worksheet</i></button>)}</div>}
        </div>}
      </section>

      <section className="automation-tool raci-tool" aria-labelledby="raci-title">
        <div className="tool-heading"><div><p className="tool-kicker">RACI integrity</p><h3 id="raci-title">Accountability check</h3></div>{raci.data && <span className={raci.data.issues.length ? "gap-count" : "clear-count"}>{raci.data.issues.length}</span>}</div>
        <p className="tool-copy">Verify that every saved RACI activity has exactly one Accountable assignment.</p>
        <button type="button" className="primary wide-action" disabled={raci.isPending} onClick={() => raci.mutate()}>{raci.isPending ? "Validating RACI…" : raci.data ? "Run validation again" : "Validate RACI matrix"}</button>
        {raci.isError && <p className="error result-message">RACI validation could not run. Try again.</p>}
        {raci.data && <div className="trace-results" aria-live="polite"><div className="check-summary"><strong>{raci.data.message}</strong><span>{raci.data.activitiesChecked} activities checked</span><small>Checked {new Date(raci.data.checkedAt).toLocaleString()}</small></div>{raci.data.issues.length ? <div className="gap-list">{raci.data.issues.map((issue) => <button key={issue.rowId} type="button" className="gap-item" onClick={() => onNavigate("PLAN", "raci")}><span className="gap-type raci">RACI</span><strong>{issue.activity}</strong><small>{issue.issue === "no one accountable" ? "No one accountable: assign exactly one A." : `Conflicting accountability: ${issue.accountableCount} people are marked A.`}</small><i>Open matrix</i></button>)}</div> : <div className="clear-state"><b>Accountability is clear</b><span>Every activity has exactly one Accountable.</span></div>}</div>}
      </section>
    </div>

    <section className="automation-tool schedule-tool" aria-labelledby="schedule-title">
      <div className="tool-heading"><div><p className="tool-kicker">Schedule health</p><h3 id="schedule-title">Automatic risk flags</h3></div>{schedule.data && <span className={schedule.data.risks.length ? "gap-count" : "clear-count"}>{schedule.data.risks.length}</span>}</div>
      <p className="tool-copy">Scan overdue work, near-term unstarted tasks, missed PERT dates, and missing ownership using today’s date.</p>
      <button type="button" className="primary wide-action" disabled={schedule.isPending} onClick={() => schedule.mutate()}>{schedule.isPending ? "Scanning schedule…" : schedule.data ? "Scan schedule again" : "Run schedule health scan"}</button>
      {schedule.isError && <p className="error result-message">The schedule scan could not run. Try again.</p>}
      {schedule.data && <div className="trace-results" aria-live="polite"><div className="check-summary"><strong>{schedule.data.message}</strong><span>{schedule.data.tasksChecked} tasks · {schedule.data.timelineItemsChecked} timeline items</span><small>Checked {new Date(schedule.data.checkedAt).toLocaleString()}</small></div>{schedule.data.risks.length ? <div className="risk-list">{schedule.data.risks.map((risk, index) => <button key={`${risk.moduleKey}-${risk.rowId}-${index}`} type="button" className="risk-item" onClick={() => onNavigate(risk.moduleKey === "task-tracker" ? "EXECUTE" : "PLAN", risk.moduleKey)}><strong>{risk.taskName}</strong><span>{risk.reason}</span><small>{risk.relevantDate === "No date" ? "No relevant date entered" : `Relevant date: ${risk.relevantDate}`}</small><i>Open worksheet</i></button>)}</div> : <div className="clear-state"><b>Schedule is clear</b><span>No defined schedule risks were found in saved task and timeline data.</span></div>}</div>}
    </section>

    <section className="automation-tool budget-tool" aria-labelledby="budget-title">
      <div className="tool-heading"><div><p className="tool-kicker">Work package costs</p><h3 id="budget-title">Budget summary</h3></div>{budgetSummary.length > 0 && <span className={budgetSummary.some((item) => item.variance > 0) ? "gap-count" : "clear-count"}>{budgetSummary.filter((item) => item.variance > 0).length}</span>}</div>
      <p className="tool-copy">Compare budgeted and actual cost for every WBS package with saved cost data.</p>
      {!budgetSummary.length ? <div className="honest-empty compact"><h3>No costs entered yet</h3><p>Add budgeted or actual cost to a WBS item to activate this summary.</p><button type="button" className="secondary inline-nav" onClick={() => onNavigate("PLAN", "wbs")}>Open WBS Builder</button></div> : <><div className="budget-totals"><div><span>Total budgeted</span><b>{money(totalBudget)}</b></div><div><span>Total actual</span><b>{money(totalActual)}</b></div><div className={totalVariance > 0 ? "over" : totalVariance < 0 ? "under" : "even"}><span>Total variance</span><b>{money(totalVariance)}</b><small>{totalVariance > 0 ? "Over budget" : totalVariance < 0 ? "Under budget" : "On budget"}</small></div></div><div className="budget-list">{budgetSummary.map((item) => <button key={item.rowId} type="button" className={item.variance > 0 ? "budget-row over" : "budget-row"} onClick={() => onNavigate("PLAN", "wbs")}><strong>{item.name}</strong><span>{money(item.actual)} actual / {money(item.budget)} budgeted</span><b>{money(item.variance)} · {item.variance > 0 ? "Over" : item.variance < 0 ? "Under" : "On budget"}</b></button>)}</div></>}
    </section>

    <section className="automation-tool report-tool" aria-labelledby="report-title">
      <div className="tool-heading"><div><p className="tool-kicker">Weekly update</p><h3 id="report-title">Status report compiler</h3></div></div>
      <p className="tool-copy">Compile completed work, schedule risk, blockers, owners, and overall progress into an editable first draft.</p>
      <div className="report-controls"><label className="field"><span>Report date</span><input aria-label="Report date" type="date" value={reportDate} onChange={(event) => { setReportDate(event.target.value); setCopyState("idle"); }} /></label><button type="button" className="primary" disabled={compile.isPending || !reportDate} onClick={() => compile.mutate()}>{compile.isPending ? "Compiling report…" : draft ? "Recompile from live data" : "Compile weekly report"}</button></div>
      {compile.isError && <p className="error result-message">The report could not be compiled. Try again.</p>}
      {compile.data && <p className="success result-message" aria-live="polite">{compile.data.message}</p>}
      {draft && <div className="draft-editor"><label className="field"><span>Editable report draft</span><textarea aria-label="Editable report draft" value={draft} onChange={(event) => { setDraft(event.target.value); setCopyState("idle"); }} /></label><div className="action-row"><button type="button" className="secondary" disabled={copyState === "copying"} onClick={() => void copyDraft()}>{copyState === "copying" ? "Copying…" : copyState === "copied" ? "Copied to clipboard" : "Copy report"}</button>{copyState === "copied" && <p className="success">Report copied.</p>}{copyState === "error" && <p className="error">Copy failed. Select the draft and copy it manually.</p>}</div></div>}
    </section>
  </div>;
}

type DonutSlice = { label: string; value: number; color: string };

function DonutChart({ slices, centerTop, centerBottom, ariaLabel }: { slices: DonutSlice[]; centerTop: string; centerBottom: string; ariaLabel: string }) {
  const positive = slices.filter((slice) => slice.value > 0);
  const total = positive.reduce((sum, slice) => sum + slice.value, 0);
  let offset = 0;
  return <div className="donut-wrap">
    <svg className="donut-chart" viewBox="0 0 120 120" role="img" aria-label={ariaLabel}>
      <circle className="donut-track" cx="60" cy="60" r="45" fill="none" strokeWidth="18" />
      {positive.map((slice) => {
        const length = slice.value / total * 100;
        const currentOffset = offset;
        offset += length;
        return <circle key={slice.label} cx="60" cy="60" r="45" fill="none" stroke={slice.color} strokeWidth="18" pathLength="100" strokeDasharray={`${length} ${100 - length}`} strokeDashoffset={-currentOffset} transform="rotate(-90 60 60)"><title>{slice.label}: {slice.value}</title></circle>;
      })}
      <text x="60" y="56" textAnchor="middle" className="donut-center-top">{centerTop}</text>
      <text x="60" y="72" textAnchor="middle" className="donut-center-bottom">{centerBottom}</text>
    </svg>
  </div>;
}

function ProgressRing({ percent, label }: { percent: number; label: string }) {
  const bounded = Math.max(0, Math.min(100, percent));
  return <svg className="progress-ring" viewBox="0 0 100 100" role="img" aria-label={`${label}: ${bounded}% complete`}>
    <circle className="ring-track" cx="50" cy="50" r="39" fill="none" strokeWidth="10" />
    <circle className="ring-value" cx="50" cy="50" r="39" fill="none" strokeWidth="10" pathLength="100" strokeDasharray={`${bounded} ${100 - bounded}`} transform="rotate(-90 50 50)" />
    <text x="50" y="55" textAnchor="middle">{bounded}%</text>
  </svg>;
}

function ChartLegend({ slices, formatValue }: { slices: DonutSlice[]; formatValue?: (value: number) => string }) {
  return <div className="chart-legend">{slices.map((slice) => <div key={slice.label}><i style={{ background: slice.color }} /><span>{slice.label}</span><b>{formatValue ? formatValue(slice.value) : slice.value}</b></div>)}</div>;
}

function ExecutiveDashboard({ onDrill }: { onDrill: (projectId: number, phase: Phase, moduleKey: string) => void }) {
  const asOfDate = localDateValue();
  const [refreshMessage, setRefreshMessage] = useState("");
  const portfolio = useQuery({ queryKey: ["portfolio-dashboard", asOfDate], queryFn: () => api.getPortfolioDashboard({ asOfDate }), refetchInterval: 15_000, refetchIntervalInBackground: true });
  const data = portfolio.data;
  const money = (value: number) => new Intl.NumberFormat(undefined, { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(value);
  const colors = ["#e9682c", "#2f7d5c", "#2c6f91", "#c98917", "#735c9e", "#a43b32", "#4d7c76", "#8b6b42"];
  const targetLabel = (days: number | null, date: string) => {
    if (days === null || !date) return "No target date";
    if (days < 0) return `${Math.abs(days)} ${Math.abs(days) === 1 ? "day" : "days"} overdue`;
    if (days === 0) return "Target due today";
    return `${days} ${days === 1 ? "day" : "days"} remaining`;
  };
  const refresh = async () => {
    setRefreshMessage("");
    const result = await portfolio.refetch();
    setRefreshMessage(result.isError ? "Refresh failed. Saved data is unchanged." : "Dashboard refreshed.");
    window.setTimeout(() => setRefreshMessage(""), 4000);
  };
  if (portfolio.isPending) return <div className="executive-loading" aria-live="polite"><span className="spinner" /><div><strong>Building portfolio view</strong><p>Reading every saved project…</p></div></div>;
  if (portfolio.isError || !data) return <div className="honest-empty"><h3>Portfolio data could not be loaded</h3><p>Your project records were not changed. Try the dashboard again.</p><button type="button" className="primary inline-nav" onClick={() => void portfolio.refetch()}>{portfolio.isFetching ? "Retrying…" : "Retry dashboard"}</button></div>;
  if (!data.projects.length) return <div className="honest-empty"><h3>No projects yet</h3><p>Create a project to activate the dashboard.</p></div>;
  const budgetItems = data.budgets.filter((item) => item.hasBudgetData && item.budgeted > 0);
  const hasBudget = data.budgets.some((item) => item.hasBudgetData);
  const allocationSlices: DonutSlice[] = budgetItems.map((item, index) => ({ label: item.projectName, value: item.budgeted, color: colors[index % colors.length] ?? colors[0] ?? "#e9682c" }));
  const taskSlices: DonutSlice[] = [
    { label: "Not Started", value: data.totals.taskStatus.notStarted, color: "#73808a" },
    { label: "In Progress", value: data.totals.taskStatus.inProgress, color: "#2c6f91" },
    { label: "Complete", value: data.totals.taskStatus.complete, color: "#2f7d5c" },
    { label: "Blocked", value: data.totals.taskStatus.blocked, color: "#a43b32" },
  ];
  const taskTotal = taskSlices.reduce((sum, slice) => sum + slice.value, 0);
  const openProject = (projectId: number) => onDrill(projectId, "MONITOR", "dashboard");
  const totalAttention = data.projects.reduce((sum, project) => sum + project.attention.traceability + project.attention.raci + project.attention.schedule, 0);
  const blockerCount = data.risks.filter((risk) => risk.kind === "Blocker").length;
  const highRiskCount = data.risks.length - blockerCount;
  return <div className="executive-dashboard">
    <section className="portfolio-pulse" aria-labelledby="portfolio-pulse-title">
      <div className="pulse-copy"><p>Portfolio pulse</p><h2 id="portfolio-pulse-title">{data.totals.projects} active {data.totals.projects === 1 ? "project" : "projects"}</h2><span>As of {data.asOfDate} · auto-refreshes every 15 seconds · updated {new Date(data.refreshedAt).toLocaleTimeString([], { hour: "numeric", minute: "2-digit", second: "2-digit" })}</span></div>
      <div className="rag-totals" aria-label="Portfolio health totals"><div className="red"><b>{data.totals.red}</b><span>Red</span></div><div className="amber"><b>{data.totals.amber}</b><span>Amber</span></div><div className="green"><b>{data.totals.green}</b><span>Green</span></div></div>
      <div className="refresh-cluster"><button type="button" className="refresh-button" disabled={portfolio.isFetching} onClick={() => void refresh()}>{portfolio.isFetching ? "Refreshing…" : "Refresh dashboard"}</button><span className={refreshMessage.includes("failed") ? "refresh-error" : "refresh-success"} aria-live="polite">{refreshMessage}</span></div>
    </section>

    <section className="visual-grid" aria-label="Portfolio charts">
      <article className="chart-panel"><div className="chart-heading"><div><p>Budget allocation</p><h3>Share by project</h3></div><b>{money(data.totals.budgeted)}</b></div>{!allocationSlices.length ? <div className="chart-empty"><strong>No budget data yet</strong><span>Add budgeted cost to a WBS package.</span></div> : <div className="donut-layout"><DonutChart slices={allocationSlices} centerTop={money(data.totals.budgeted)} centerBottom="allocated" ariaLabel="Budget allocation across projects" /><ChartLegend slices={allocationSlices} formatValue={money} /></div>}</article>
      <article className="chart-panel"><div className="chart-heading"><div><p>Task status</p><h3>Portfolio breakdown</h3></div><b>{taskTotal} tasks</b></div>{!taskTotal ? <div className="chart-empty"><strong>No task data yet</strong><span>Add tasks or promote WBS packages to activate this chart.</span></div> : <div className="donut-layout"><DonutChart slices={taskSlices} centerTop={String(taskTotal)} centerBottom="tasks" ariaLabel="Task status breakdown across the portfolio" /><ChartLegend slices={taskSlices} /></div>}</article>
    </section>

    <section className="executive-section" aria-labelledby="budget-bars-title">
      <div className="executive-section-heading"><div><p>Budget performance</p><h3 id="budget-bars-title">Budget vs actual</h3></div><span>Actual turns red when it exceeds budget</span></div>
      {!hasBudget ? <div className="chart-empty wide"><strong>No budget data yet</strong><span>Add budgeted or actual cost to a WBS package.</span></div> : <><div className="executive-budget-totals"><div><span>Total budgeted</span><b>{money(data.totals.budgeted)}</b></div><div><span>Total actual</span><b>{money(data.totals.actual)}</b></div><div className={data.totals.variance > 0 ? "over" : data.totals.variance < 0 ? "under" : "even"}><span>Total variance</span><b>{money(data.totals.variance)}</b><small>{data.totals.variance > 0 ? "Over budget" : data.totals.variance < 0 ? "Under budget" : "On budget"}</small></div></div><div className="comparison-chart">{data.budgets.filter((item) => item.hasBudgetData).map((item) => { const scale = Math.max(item.budgeted, item.actual, 1); return <button key={item.projectId} type="button" className="comparison-row" onClick={() => onDrill(item.projectId, "PLAN", "wbs")}><div className="comparison-label"><strong>{item.projectName}</strong><span>{item.projectCode}</span></div><div className="bar-pair"><span className="bar-line"><i className="bar-budget" style={{ width: `${Math.min(100, item.budgeted / scale * 100)}%` }} /><em>{money(item.budgeted)}</em></span><span className="bar-line"><i className={item.variance > 0 ? "bar-actual over" : "bar-actual"} style={{ width: `${Math.min(100, item.actual / scale * 100)}%` }} /><em>{money(item.actual)}</em></span></div></button>; })}</div><div className="bar-legend"><span><i className="budget" />Budget</span><span><i className="actual" />Actual</span><span><i className="over" />Over budget actual</span></div></>}
    </section>

    <section className="executive-section" aria-labelledby="completion-title"><div className="executive-section-heading"><div><p>Delivery</p><h3 id="completion-title">Project completion</h3></div><span>Live from task progress</span></div><div className="completion-rings">{data.projects.map((project) => <button type="button" key={project.projectId} className="completion-card" onClick={() => openProject(project.projectId)}><ProgressRing percent={project.percentComplete} label={project.projectName} /><span><strong>{project.projectName}</strong><small>{project.taskCount ? `${project.taskCount} tracked ${project.taskCount === 1 ? "task" : "tasks"}` : "No task data yet"}</small></span></button>)}</div></section>

    <section className="executive-section" aria-labelledby="health-title">
      <div className="executive-section-heading"><div><p>Explainable RAG</p><h3 id="health-title">Project health</h3></div><span>Open a card for project detail</span></div>
      <div className="rag-health-grid">{data.projects.map((project) => <button type="button" key={project.projectId} className={`rag-health-card health-${project.health.toLowerCase()}`} onClick={() => openProject(project.projectId)}>
        <span className="rag-label">{project.health}</span><div><small>{project.projectCode}</small><h4>{project.projectName}</h4><p>{project.reasons.join(" · ")}</p></div><footer><span>{project.percentComplete}% complete</span><b>{targetLabel(project.daysToTarget, project.targetDate)}</b></footer>
      </button>)}</div>
    </section>

    <section className="executive-section" aria-labelledby="dates-title"><div className="executive-section-heading"><div><p>Next 14 days</p><h3 id="dates-title">Key date timeline</h3></div><span>{data.dates.length} items, overdue included</span></div>{!data.dates.length ? <div className="quiet-empty"><b>No deadlines in view</b><span>No overdue items or upcoming dates in the next 14 days.</span></div> : <div className="timeline-strip">{data.dates.map((item, index) => <button key={`${item.projectId}-${item.type}-${item.date}-${index}`} type="button" className={item.overdue ? "timeline-event overdue" : "timeline-event"} onClick={() => onDrill(item.projectId, item.type === "Task" || item.type === "Key date" ? "EXECUTE" : "PLAN", item.type === "Task" ? "task-tracker" : item.type === "Key date" ? "calendar" : item.type === "Milestone" ? "timeline" : "wbs")}><span className="timeline-date"><b>{item.date.slice(5)}</b><small>{item.overdue ? "Overdue" : item.type}</small></span><i /><div><strong>{item.title}</strong><span>{item.projectName}</span><small>{item.owner}</small></div></button>)}</div>}</section>

    <section className="summary-tiles" aria-label="Portfolio attention counts">
      <button type="button" className="summary-tile blockers" onClick={() => { const item = data.risks.find((risk) => risk.kind === "Blocker"); if (item) onDrill(item.projectId, "EXECUTE", "task-tracker"); }} disabled={!blockerCount}><span>Blockers</span><b>{blockerCount}</b><small>{blockerCount ? "Open blocker details" : "None open"}</small></button>
      <button type="button" className="summary-tile risks" onClick={() => { const item = data.risks.find((risk) => risk.kind !== "Blocker"); if (item) onDrill(item.projectId, "PLAN", "ac-log"); }} disabled={!highRiskCount}><span>High-impact risks</span><b>{highRiskCount}</b><small>{highRiskCount ? "Open assumptions and constraints" : "None open"}</small></button>
      <button type="button" className="summary-tile attention" onClick={() => { const project = data.projects.find((item) => item.attention.traceability + item.attention.raci + item.attention.schedule > 0); if (project) onDrill(project.projectId, "MONITOR", "automation"); }} disabled={!totalAttention}><span>Attention flags</span><b>{totalAttention}</b><small>{totalAttention ? "Traceability, RACI, and schedule" : "All checks clear"}</small></button>
    </section>

    <div className="executive-columns">
      <section className="executive-section" aria-labelledby="risks-title"><div className="executive-section-heading"><div><p>Escalations</p><h3 id="risks-title">Blockers & risks</h3></div><span>{data.risks.length} open</span></div>{!data.risks.length ? <div className="quiet-empty"><b>No open escalations</b><span>No task blockers or high-impact assumptions and constraints are recorded.</span></div> : <div className="executive-list risks-list">{data.risks.map((item, index) => <button key={`${item.projectId}-${item.kind}-${index}`} type="button" className="risk-row" onClick={() => onDrill(item.projectId, item.kind === "Blocker" ? "EXECUTE" : "PLAN", item.kind === "Blocker" ? "task-tracker" : "ac-log")}><span className={`risk-kind ${item.kind.toLowerCase()}`}>{item.kind}</span><div><strong>{item.title}</strong><p>{item.detail}</p><small>{item.projectName} · {item.owner} · {item.impact}</small></div></button>)}</div>}</section>
      <section className="executive-section" aria-labelledby="attention-title"><div className="executive-section-heading"><div><p>Automation</p><h3 id="attention-title">Attention by project</h3></div><span>Tap a tile to investigate</span></div><div className="attention-project-grid">{data.projects.map((project) => { const total = project.attention.traceability + project.attention.raci + project.attention.schedule; return <button key={project.projectId} type="button" className="attention-project" onClick={() => onDrill(project.projectId, "MONITOR", "automation")}><span><strong>{project.projectName}</strong><small>{project.projectCode}</small></span><em className={total ? "needs-attention" : "clear"}>{total}</em><div><span>Trace <b>{project.attention.traceability}</b></span><span>RACI <b>{project.attention.raci}</b></span><span>Schedule <b>{project.attention.schedule}</b></span></div></button>; })}</div></section>
    </div>
  </div>;
}

function Dashboard({ rows }: { rows: DataRow[] }) {
  const tasks = rows.filter((r) => r.moduleKey === "task-tracker").map((r) => parseJson(r.dataJson));
  const timeline = rows.filter((r) => r.moduleKey === "timeline").map((r) => parseJson(r.dataJson));
  const overall = tasks.length ? Math.round(tasks.reduce((sum, t) => sum + Number(t.percentComplete || 0), 0) / tasks.length) : 0;
  const counts = statusOptions.map((status) => ({ status, count: tasks.filter((t) => t.status === status).length }));
  const timelineProgress = timeline.length ? Math.round(timeline.reduce((sum, t) => sum + Number(t.percentComplete || 0), 0) / timeline.length) : 0;
  return <div className="dashboard">
    {!tasks.length ? <div className="honest-empty"><h3>No task data yet</h3><p>Add tasks in Execute, or promote WBS items, to activate project health.</p></div> : <>
      <section className="health-hero"><div className="health-number">{overall}<span>%</span></div><div><p>Overall completion</p><h3>{overall >= 80 ? "Closing strong" : overall >= 40 ? "Work is moving" : "Early delivery"}</h3></div></section>
      <div className="metric-list">{counts.map((c) => <div key={c.status}><span>{c.status}</span><b>{c.count}</b></div>)}</div>
      <section className="progress-block"><div><span>Timeline progress</span><b>{timelineProgress}%</b></div><div className="progress-track"><i style={{ width: `${timelineProgress}%` }} /></div><p>{timeline.length} timeline {timeline.length === 1 ? "item" : "items"}</p></section>
    </>}
  </div>;
}

export function App() {
  const projectsQuery = useQuery({ queryKey: ["projects"], queryFn: () => api.listProjects({}) });
  const [selectedId, setSelectedId] = useState<number>(0);
  const [topLevel, setTopLevel] = useState<"DASHBOARD" | Phase>("DASHBOARD");
  const [moduleKey, setModuleKey] = useState("business-case");
  const projects = projectsQuery.data?.projects ?? [];
  const selected = projects.find((p) => p.id === selectedId) ?? projects[0];
  const phase: Phase = topLevel === "DASHBOARD" ? "MONITOR" : topLevel;
  const projectData = useQuery({ queryKey: ["project", selected?.id], queryFn: () => api.getProjectData({ projectId: selected?.id ?? 0 }), enabled: Boolean(selected?.id) && topLevel !== "DASHBOARD" });
  const modules = moduleDefs.filter((m) => m.phase === phase);
  const activeModule = modules.find((m) => m.key === moduleKey) ?? modules[0] ?? moduleDefs[0];
  const data = projectData.data;
  const section = data?.sections.find((s) => s.moduleKey === activeModule?.key);
  const rows = data?.rows.filter((r) => r.moduleKey === activeModule?.key) ?? [];
  const raciRoles = useMemo(() => parseJson(section?.dataJson).roles?.split(",").map((s) => s.trim()).filter(Boolean).slice(0, 8) ?? [], [section?.dataJson]);

  if (projectsQuery.isPending) return <main className="loading"><SafeAreaTopScrim backgroundColor="var(--bg)" /><p>Loading project workspace…</p></main>;
  if (projectsQuery.isError) return <main className="loading"><SafeAreaTopScrim backgroundColor="var(--bg)" /><p>Projects could not be loaded. Try opening the artifact again.</p></main>;
  if (!projects.length) return <main className="shell empty-shell"><SafeAreaTopScrim backgroundColor="var(--bg)" /><CreateProject onCreated={setSelectedId} /></main>;

  const changeTopLevel = (next: "DASHBOARD" | Phase) => {
    setTopLevel(next);
    if (next !== "DASHBOARD") {
      const first = moduleDefs.find((m) => m.phase === next);
      if (first) setModuleKey(first.key);
    }
  };
  const drill = (projectId: number, nextPhase: Phase, nextModule: string) => {
    setSelectedId(projectId);
    setTopLevel(nextPhase);
    setModuleKey(nextModule);
  };
  const navItems: Array<{ key: "DASHBOARD" | Phase; label: string }> = [
    { key: "DASHBOARD", label: "Dashboard" },
    ...phases.map((item) => ({ key: item, label: phaseNames[item] })),
  ];
  return <div className="app"><SafeAreaTopScrim backgroundColor="var(--bg)" />
    <header className="topbar"><div><label htmlFor="project-switcher">Active project</label><select id="project-switcher" value={selected?.id ?? ""} onChange={(e) => setSelectedId(Number(e.target.value))}>{projects.map((p) => <option key={p.id} value={p.id}>{p.name} · {p.projectCode}</option>)}</select></div><details className="new-project"><summary>New project</summary><CreateProject compact onCreated={setSelectedId} /></details></header>
    <main className={`workspace ${topLevel === "DASHBOARD" ? "dashboard-workspace" : ""}`}>
      <nav className="phase-route top-level-route" aria-label="Platform navigation">{navItems.map((item, index) => <button key={item.key} className={topLevel === item.key ? "active" : ""} onClick={() => changeTopLevel(item.key)}><span>{item.key === "DASHBOARD" ? "●" : String(index).padStart(2, "0")}</span>{item.label}</button>)}</nav>
      {selected && <ExcelTools selected={selected} onImported={(id) => { setSelectedId(id); setTopLevel("INITIATE"); setModuleKey("business-case"); }} />}
      {topLevel === "DASHBOARD" ? <section className="dashboard-stage"><div className="dashboard-stage-heading"><div><p>Executive view</p><h1>Dashboard</h1></div><span>Live portfolio data</span></div><ExecutiveDashboard onDrill={drill} /></section> : <>
        {selected && <ProjectHeaderForm key={selected.id} project={selected} />}
        <div className="work-area">
          <aside className="step-nav"><p>{phaseNames[phase]} flow</p>{modules.map((m, i) => <button key={m.key} className={activeModule?.key === m.key ? "active" : ""} onClick={() => setModuleKey(m.key)}><span>{String(i + 1).padStart(2, "0")}</span>{m.label}</button>)}</aside>
          <section className="module-panel">
            <label className="mobile-step">Current worksheet<select aria-label="Current worksheet" value={activeModule?.key} onChange={(e) => setModuleKey(e.target.value)}>{modules.map((m) => <option key={m.key} value={m.key}>{m.label}</option>)}</select></label>
            {projectData.isPending ? <p className="loading-inline">Loading project data…</p> : projectData.isError ? <div className="honest-empty"><h3>Project data could not be loaded</h3><p>Switch projects or reopen the artifact and try again.</p></div> : activeModule && <>
              <div className="module-heading"><p>{phaseNames[phase]}</p><h2>{activeModule.label}</h2><span>{activeModule.intro}</span></div>
              {activeModule.key === "dashboard" ? <Dashboard rows={data?.rows ?? []} /> : activeModule.key === "automation" ? <AutomationCenter key={selected?.id ?? 0} projectId={selected?.id ?? 0} rows={data?.rows ?? []} onNavigate={(nextPhase, nextModule) => { setTopLevel(nextPhase); setModuleKey(nextModule); }} /> : <>
                <SectionForm key={`${selected?.id ?? 0}-${activeModule.key}`} projectId={selected?.id ?? 0} module={activeModule} dataJson={section?.dataJson} updatedAt={section?.updatedAt} />
                {activeModule.key === "raci" && !raciRoles.length && <p className="callout">Save at least one role above before adding matrix activities.</p>}
                <div className="rows">{rows.map((r) => <RowEditor key={r.id} row={r} projectId={selected?.id ?? 0} module={activeModule} raciRoles={raciRoles} allRows={data?.rows ?? []} />)}</div>
                {activeModule.rowFields && !rows.length && <div className="honest-empty compact"><h3>No {activeModule.rowLabel?.toLowerCase()}s yet</h3><p>Add the first entry when the project is ready.</p></div>}
                {(activeModule.key !== "raci" || raciRoles.length > 0) && <AddRow projectId={selected?.id ?? 0} module={activeModule} raciRoles={raciRoles} />}
              </>}
            </>}
          </section>
        </div>
      </>}
    </main>
  </div>;
}
