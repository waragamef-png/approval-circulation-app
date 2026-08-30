import assert from "node:assert/strict";
import test from "node:test";
import { createGraphSharePoint, isSharePointTargetConfigured } from "./graphSharePoint.mjs";

test("対象サイトの文書ライブラリから対応ファイルを一覧取得する", async () => {
  const requested = [];
  let now = 1_000;
  const responses = [
    { id: "site-1", displayName: "承認回覧", webUrl: "https://example.sharepoint.com/sites/approval" },
    { value: [{ id: "drive-1", name: "共有ドキュメント", driveType: "documentLibrary" }] },
    {
      value: [
        { id: "word-1", name: "手順書.docx", file: {}, webUrl: "https://example/word", lastModifiedDateTime: "2026-08-01T00:00:00Z", lastModifiedBy: { user: { displayName: "佐藤 花子" } } },
        { id: "folder-1", name: "保管", folder: {}, webUrl: "https://example/folder" },
        { id: "pdf-1", name: "仕様書.pdf", file: {}, webUrl: "https://example/pdf", lastModifiedDateTime: "2026-08-02T00:00:00Z", lastModifiedBy: { application: { displayName: "自動更新" } } },
        { id: "text-1", name: "メモ.txt", file: {}, webUrl: "https://example/text" },
      ],
      "@odata.nextLink": "https://graph.example/next-page",
    },
    {
      value: [
        { id: "excel-1", name: "一覧.xlsx", file: {}, webUrl: "https://example/excel", lastModifiedDateTime: "2026-08-03T00:00:00Z", lastModifiedBy: {} },
      ],
    },
  ];
  const sharePoint = createGraphSharePoint({
    acquireAccessToken: async () => "test-token",
    hostname: "example.sharepoint.com",
    sitePath: "/sites/approval",
    folderPath: "回覧対象/2026年度",
    fetchImpl: async (url, options) => {
      requested.push({ url, authorization: options.headers.Authorization });
      return { ok: true, json: async () => responses.shift(), headers: { get: () => null } };
    },
    graphBaseUrl: "https://graph.example/v1.0",
    now: () => now,
  });

  const first = await sharePoint.listFiles();
  const second = await sharePoint.listFiles();

  assert.equal(first, second, "2回目はキャッシュを使用する");
  assert.deepEqual(first.files.map((file) => file.id), ["pdf-1", "excel-1", "word-1"]);
  assert.equal(first.files[0].updatedBy, "自動更新");
  assert.equal(first.files[1].updatedBy, "更新者不明");
  assert.equal(first.files[0].location, "承認回覧 / 共有ドキュメント / 回覧対象/2026年度");
  assert.equal(first.libraryName, "共有ドキュメント");
  assert.equal(requested.length, 4);
  assert.ok(requested[0].url.includes("/sites/example.sharepoint.com:/sites/approval"));
  assert.ok(requested[2].url.includes("root:/%E5%9B%9E%E8%A6%A7%E5%AF%BE%E8%B1%A1/2026%E5%B9%B4%E5%BA%A6:/children"));
  assert.ok(requested.every((request) => request.authorization === "Bearer test-token"));

  now += 60_001;
  sharePoint.clearCache();
});

test("SharePoint参照先の設定有無を判定する", () => {
  assert.equal(isSharePointTargetConfigured({ SHAREPOINT_HOSTNAME: "example.sharepoint.com", SHAREPOINT_SITE_PATH: "/sites/approval" }), true);
  assert.equal(isSharePointTargetConfigured({ SHAREPOINT_HOSTNAME: "example.sharepoint.com" }), false);
});

test("Graphエラーには確認用リクエストIDを含める", async () => {
  const sharePoint = createGraphSharePoint({
    acquireAccessToken: async () => "test-token",
    hostname: "example.sharepoint.com",
    sitePath: "/sites/approval",
    fetchImpl: async () => ({ ok: false, status: 403, headers: { get: (name) => name === "request-id" ? "request-123" : null } }),
  });
  await assert.rejects(() => sharePoint.listFiles(), /403.*request-id=request-123/);
});
