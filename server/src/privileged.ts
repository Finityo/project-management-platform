import { execFile } from "node:child_process";
import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import { definePrivilegedContracts, definePrivilegedHandlers, z } from "@hatch/space-sdk";

const execFileAsync = promisify(execFile);
const PAGES_PER_RENDER = 3;

type CommandError = { code?: unknown; message?: unknown; stderr?: unknown };

function commandDetails(error: unknown): { code: string; message: string; stderr: string } {
  const value = error && typeof error === "object" ? error as CommandError : {};
  return {
    code: typeof value.code === "string" ? value.code : "",
    message: typeof value.message === "string" ? value.message : "",
    stderr: typeof value.stderr === "string" ? value.stderr : Buffer.isBuffer(value.stderr) ? value.stderr.toString("utf8") : "",
  };
}

function isPasswordFailure(details: { message: string; stderr: string }): boolean {
  return /incorrect password|password required|encrypted.*password/i.test(`${details.message}\n${details.stderr}`);
}

export const privileged = definePrivilegedContracts({
  renderPdfPages: {
    request: z.object({ dataBase64: z.string().max(140_000_000), firstPage: z.number().int().positive() }),
    response: z.object({
      ok: z.boolean(),
      errorCode: z.enum(["encrypted", "invalid_pdf", "renderer_unavailable", "render_failed"]).optional(),
      pageCount: z.number().int().nonnegative(),
      pages: z.array(z.object({ dataBase64: z.string(), filename: z.string(), pageNumber: z.number().int().positive() })).max(3),
    }),
    timeoutMs: 90_000,
  },
});

export const privilegedHandlers = definePrivilegedHandlers(privileged, {
  async renderPdfPages(args) {
    const dir = await mkdtemp(join(tmpdir(), "drawing-scan-"));
    try {
      const input = join(dir, "drawing.pdf");
      const outputPrefix = join(dir, "page");
      await writeFile(input, Buffer.from(args.dataBase64, "base64"));

      let pageCount = 0;
      try {
        const info = await execFileAsync("/usr/bin/pdfinfo", [input], { timeout: 20_000, maxBuffer: 8_000_000 });
        const match = String(info.stdout).match(/^Pages:\s+(\d+)\s*$/m);
        pageCount = Number(match?.[1] ?? 0);
        if (!pageCount) return { ok: false, errorCode: "invalid_pdf" as const, pageCount: 0, pages: [] };
      } catch (error) {
        const details = commandDetails(error);
        console.error("PDF inspection failed", { code: details.code, stderr: details.stderr.slice(0, 2_000) });
        if (isPasswordFailure(details)) return { ok: false, errorCode: "encrypted" as const, pageCount: 0, pages: [] };
        if (details.code === "ENOENT") return { ok: false, errorCode: "renderer_unavailable" as const, pageCount: 0, pages: [] };
        return { ok: false, errorCode: "invalid_pdf" as const, pageCount: 0, pages: [] };
      }

      if (args.firstPage > pageCount) return { ok: false, errorCode: "render_failed" as const, pageCount, pages: [] };
      const lastPage = Math.min(pageCount, args.firstPage + PAGES_PER_RENDER - 1);
      try {
        await execFileAsync("/usr/bin/pdftoppm", [
          "-f", String(args.firstPage), "-l", String(lastPage), "-r", "110", "-jpeg", "-jpegopt", "quality=75", input, outputPrefix,
        ], { timeout: 60_000, maxBuffer: 8_000_000 });
      } catch (error) {
        const details = commandDetails(error);
        console.error("PDF page rendering failed", { code: details.code, firstPage: args.firstPage, lastPage, stderr: details.stderr.slice(0, 2_000) });
        if (isPasswordFailure(details)) return { ok: false, errorCode: "encrypted" as const, pageCount, pages: [] };
        if (details.code === "ENOENT") return { ok: false, errorCode: "renderer_unavailable" as const, pageCount, pages: [] };
        return { ok: false, errorCode: "render_failed" as const, pageCount, pages: [] };
      }

      const names = (await readdir(dir)).filter((name) => /^page-\d+\.jpg$/.test(name)).sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).slice(0, PAGES_PER_RENDER);
      const pages = await Promise.all(names.map(async (filename, index) => ({
        filename,
        pageNumber: args.firstPage + index,
        dataBase64: (await readFile(join(dir, filename))).toString("base64"),
      })));
      if (!pages.length) return { ok: false, errorCode: "render_failed" as const, pageCount, pages: [] };
      return { ok: true, pageCount, pages };
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  },
});
