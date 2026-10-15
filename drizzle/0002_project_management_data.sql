DROP TABLE entries;
--> statement-breakpoint
CREATE TABLE projects (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  project_code TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL,
  sponsor TEXT NOT NULL DEFAULT '',
  project_manager TEXT NOT NULL DEFAULT '',
  business_owner TEXT NOT NULL DEFAULT '',
  project_date TEXT NOT NULL DEFAULT '',
  version TEXT NOT NULL DEFAULT '1.0',
  overview TEXT NOT NULL DEFAULT '',
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL
);
--> statement-breakpoint
CREATE TABLE sections (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
  module_key TEXT NOT NULL,
  data_json TEXT NOT NULL DEFAULT '{}',
  updated_at INTEGER NOT NULL,
  UNIQUE(project_id, module_key)
);
--> statement-breakpoint
CREATE TABLE module_rows (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
  module_key TEXT NOT NULL,
  data_json TEXT NOT NULL DEFAULT '{}',
  sort_order INTEGER NOT NULL DEFAULT 0,
  source_row_id INTEGER,
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL
);
--> statement-breakpoint
CREATE INDEX module_rows_project_module_idx ON module_rows(project_id, module_key);
