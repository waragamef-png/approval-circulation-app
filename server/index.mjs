import "dotenv/config";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import cors from "cors";
import express from "express";
import { createAccessTokenProvider, isGraphConfigured } from "./graphAuth.mjs";
import { createGraphDirectory } from "./graphDirectory.mjs";
import { createGraphSharePoint, isSharePointTargetConfigured } from "./graphSharePoint.mjs";

const app = express();
const port = Number(process.env.APP_PORT || 3000);
const host = process.env.APP_HOST?.trim() || "127.0.0.1";
const allowedOrigins = (process.env.APP_ALLOWED_ORIGINS || "http://localhost:5173,http://127.0.0.1:5173")
  .split(",")
  .map((value) => value.trim())
  .filter(Boolean);

app.disable("x-powered-by");
app.use(cors({
  origin(origin, callback) {
    if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
    return callback(new Error("Origin is not allowed"));
  },
}));
app.use(express.json({ limit: "32kb" }));

let accessTokenProviderPromise;
function getAccessTokenProvider() {
  if (!accessTokenProviderPromise) {
    accessTokenProviderPromise = createAccessTokenProvider().catch((error) => {
      accessTokenProviderPromise = undefined;
      throw error;
    });
  }
  return accessTokenProviderPromise;
}

let directoryPromise;
async function getDirectory() {
  if (!directoryPromise) {
    directoryPromise = getAccessTokenProvider()
      .then((acquireAccessToken) => createGraphDirectory({
        acquireAccessToken,
        cacheTtlMs: Number(process.env.GRAPH_DIRECTORY_CACHE_TTL_MS || 300_000),
        maxUsers: Number(process.env.GRAPH_DIRECTORY_MAX_USERS || 5_000),
      }))
      .catch((error) => {
        directoryPromise = undefined;
        throw error;
      });
  }
  return directoryPromise;
}

let sharePointPromise;
async function getSharePoint() {
  if (!sharePointPromise) {
    sharePointPromise = getAccessTokenProvider()
      .then((acquireAccessToken) => createGraphSharePoint({
        acquireAccessToken,
        hostname: process.env.SHAREPOINT_HOSTNAME || process.env.VITE_SHAREPOINT_HOSTNAME,
        sitePath: process.env.SHAREPOINT_SITE_PATH || process.env.VITE_SHAREPOINT_SITE_PATH,
        libraryName: process.env.SHAREPOINT_LIBRARY_NAME,
        folderPath: process.env.SHAREPOINT_FOLDER_PATH,
        cacheTtlMs: Number(process.env.SHAREPOINT_CACHE_TTL_MS || 60_000),
        maxFiles: Number(process.env.SHAREPOINT_MAX_FILES || 500),
      }))
      .catch((error) => {
        sharePointPromise = undefined;
        throw error;
      });
  }
  return sharePointPromise;
}

app.get("/api/health", (_request, response) => {
  response.json({ status: "ok", graphConfigured: isGraphConfigured(), sharePointConfigured: isSharePointTargetConfigured() });
});

app.get("/api/directory/users", async (request, response) => {
  const query = String(request.query.q || "").trim();
  const limit = Math.min(Math.max(Number(request.query.limit) || 10, 1), 20);
  if (!query) return response.status(400).json({ message: "検索文字を入力してください。" });
  if (!isGraphConfigured()) return response.status(503).json({ message: "Microsoft 365連携が未設定です。" });

  try {
    const directory = await getDirectory();
    const users = await directory.search(query, limit);
    return response.json({ source: "microsoft-graph", users });
  } catch (error) {
    console.error("Directory search failed:", error instanceof Error ? error.message : error);
    return response.status(502).json({ message: "Microsoft 365の社内名簿を取得できませんでした。" });
  }
});

app.get("/api/sharepoint/files", async (_request, response) => {
  if (!isGraphConfigured()) return response.status(503).json({ message: "Microsoft 365連携が未設定です。" });
  if (!isSharePointTargetConfigured()) return response.status(503).json({ message: "SharePointの参照先が未設定です。" });

  try {
    const sharePoint = await getSharePoint();
    const result = await sharePoint.listFiles();
    return response.json({ source: "microsoft-graph", ...result });
  } catch (error) {
    console.error("SharePoint file listing failed:", error instanceof Error ? error.message : error);
    return response.status(502).json({ message: "SharePointの文書一覧を取得できませんでした。" });
  }
});

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const distDir = path.join(rootDir, "dist");
if (existsSync(path.join(distDir, "index.html"))) {
  app.use(express.static(distDir));
  app.get("/{*path}", (_request, response) => response.sendFile(path.join(distDir, "index.html")));
}

app.use((error, _request, response, _next) => {
  if (error instanceof Error && error.message === "Origin is not allowed") {
    return response.status(403).json({ message: "この接続元は許可されていません。" });
  }
  console.error(error);
  return response.status(500).json({ message: "サーバーでエラーが発生しました。" });
});

app.listen(port, host, () => {
  console.log(`社内承認回覧サーバー: http://${host}:${port}`);
  console.log(`Microsoft 365ユーザー検索: ${isGraphConfigured() ? "設定済み" : "未設定（ローカルデータを使用）"}`);
  console.log(`SharePoint文書参照: ${isGraphConfigured() && isSharePointTargetConfigured() ? "設定済み" : "未設定（確認用データを使用）"}`);
});
