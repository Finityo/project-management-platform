import { integer, sqliteTable, text, uniqueIndex, index } from "drizzle-orm/sqlite-core";

export const projects = sqliteTable("projects", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectCode: text("project_code").notNull(),
  name: text("name").notNull(),
  sponsor: text("sponsor").notNull().default(""),
  projectManager: text("project_manager").notNull().default(""),
  businessOwner: text("business_owner").notNull().default(""),
  projectDate: text("project_date").notNull().default(""),
  version: text("version").notNull().default("1.0"),
  overview: text("overview").notNull().default(""),
  createdAt: integer("created_at", { mode: "timestamp_ms" }).notNull().$defaultFn(() => new Date()),
  updatedAt: integer("updated_at", { mode: "timestamp_ms" }).notNull().$defaultFn(() => new Date()),
}, (table) => [uniqueIndex("projects_code_idx").on(table.projectCode)]);

export const sections = sqliteTable("sections", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectId: integer("project_id").notNull().references(() => projects.id, { onDelete: "cascade" }),
  moduleKey: text("module_key").notNull(),
  dataJson: text("data_json").notNull().default("{}"),
  updatedAt: integer("updated_at", { mode: "timestamp_ms" }).notNull().$defaultFn(() => new Date()),
}, (table) => [uniqueIndex("sections_project_module_idx").on(table.projectId, table.moduleKey)]);

export const moduleRows = sqliteTable("module_rows", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectId: integer("project_id").notNull().references(() => projects.id, { onDelete: "cascade" }),
  moduleKey: text("module_key").notNull(),
  dataJson: text("data_json").notNull().default("{}"),
  sortOrder: integer("sort_order").notNull().default(0),
  sourceRowId: integer("source_row_id"),
  createdAt: integer("created_at", { mode: "timestamp_ms" }).notNull().$defaultFn(() => new Date()),
  updatedAt: integer("updated_at", { mode: "timestamp_ms" }).notNull().$defaultFn(() => new Date()),
}, (table) => [index("module_rows_project_module_idx").on(table.projectId, table.moduleKey)]);
