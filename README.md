# Project Management Platform

A construction project management web app: initiate, plan, execute, and monitor projects with WBS-fed task tracking, PERT estimates, a dashboard, and Excel workbook import/export.

## Stack

- Client: React 19 + TypeScript + Tailwind CSS 4, Recharts for charts, xlsx for workbook import/export
- Server: TypeScript actions on Bun, Drizzle ORM + SQLite
- Package manager: Bun

## Run it

```sh
bun install
bun run build
```

Typecheck client and server:

```sh
bun run typecheck
```

## Notes

- `drizzle/` holds the SQL migrations.
- The app database (`app.db`) is gitignored; it is created/managed at runtime.
- Built and deployed through the web-artifact builder; `bun run build` alone does not publish the artifact.
