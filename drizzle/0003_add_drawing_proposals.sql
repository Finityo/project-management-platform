CREATE TABLE drawing_proposals (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
  file_name TEXT NOT NULL,
  mime_type TEXT NOT NULL,
  blob_key TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'draft',
  markup_basis_points INTEGER NOT NULL DEFAULT 2000,
  analysis_json TEXT NOT NULL DEFAULT '{}',
  proposal_json TEXT NOT NULL DEFAULT '{}',
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL
);
--> statement-breakpoint
CREATE INDEX drawing_proposals_project_idx ON drawing_proposals(project_id, updated_at);
