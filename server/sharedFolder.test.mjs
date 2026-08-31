import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { contentDisposition, createSharedFolderGateway, parseSharedFolderRoots, SharedFolderError } from "./sharedFolder.mjs";

test("許可ルート内のPDFを読み取り用に解決する", async (context) => {
  const base = await mkdtemp(path.join(os.tmpdir(), "approval-shared-folder-"));
  context.after(() => rm(base, { recursive: true, force: true }));
  const root = path.join(base, "allowed");
  const filePath = path.join(root, "確認資料.pdf");
  await mkdir(root);
  await writeFile(filePath, "%PDF-1.4\n%%EOF\n");

  const gateway = createSharedFolderGateway({ roots: [root] });
  const result = await gateway.resolveFile(filePath);

  assert.equal(result.absolutePath, await realPath(filePath));
  assert.equal(result.fileName, "確認資料.pdf");
  assert.equal(result.inline, true);
});

test("許可ルート外、相対パス、未対応形式を拒否する", async (context) => {
  const base = await mkdtemp(path.join(os.tmpdir(), "approval-shared-folder-"));
  context.after(() => rm(base, { recursive: true, force: true }));
  const root = path.join(base, "allowed");
  const outside = path.join(base, "outside.xlsx");
  await mkdir(root);
  await writeFile(outside, "test");

  const gateway = createSharedFolderGateway({ roots: [root] });
  await assert.rejects(() => gateway.resolveFile(outside), (error) => error instanceof SharedFolderError && error.statusCode === 403);
  await assert.rejects(() => gateway.resolveFile(path.join(base, "missing-outside.xlsx")), (error) => error instanceof SharedFolderError && error.statusCode === 403);
  await assert.rejects(() => gateway.resolveFile("relative.pdf"), (error) => error instanceof SharedFolderError && error.statusCode === 400);
  await assert.rejects(() => gateway.resolveFile(path.join(root, "script.html")), (error) => error instanceof SharedFolderError && error.statusCode === 400);
});

test("未設定と存在しないファイルを区別する", async (context) => {
  const base = await mkdtemp(path.join(os.tmpdir(), "approval-shared-folder-"));
  context.after(() => rm(base, { recursive: true, force: true }));
  const root = path.join(base, "allowed");
  await mkdir(root);

  await assert.rejects(() => createSharedFolderGateway().resolveFile(path.join(root, "missing.pdf")), (error) => error instanceof SharedFolderError && error.statusCode === 503);
  await assert.rejects(() => createSharedFolderGateway({ roots: [root] }).resolveFile(path.join(root, "missing.pdf")), (error) => error instanceof SharedFolderError && error.statusCode === 404);
});

test("ルート一覧と日本語ファイル名の応答ヘッダーを生成する", () => {
  const roots = parseSharedFolderRoots(` C:${path.sep}Docs ; D:${path.sep}Forms `);
  assert.equal(roots.length, 2);
  assert.match(contentDisposition("申請書.xlsx", false), /^attachment;/);
  assert.match(contentDisposition("申請書.xlsx", false), /filename\*=UTF-8''/);
});

async function realPath(value) {
  const { realpath } = await import("node:fs/promises");
  return realpath(value);
}
