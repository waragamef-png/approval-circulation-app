import type { MockFile } from "../types";

interface SharePointFilesResponse {
  source: "microsoft-graph";
  siteName: string;
  siteUrl: string;
  libraryName: string;
  folderPath: string;
  files: MockFile[];
}

const configuredValue = (import.meta.env.VITE_API_BASE_URL || "").trim();
const configuredBaseUrl = configuredValue === "same-origin" ? "" : configuredValue.replace(/\/$/, "");

export const sharePointFileListingEnabled = configuredValue.length > 0;

export async function listSharePointFiles(signal?: AbortSignal): Promise<SharePointFilesResponse> {
  const response = await fetch(`${configuredBaseUrl}/api/sharepoint/files`, { signal });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({})) as { message?: string };
    throw new Error(payload.message || "SharePointの文書一覧を取得できませんでした。");
  }
  return response.json() as Promise<SharePointFilesResponse>;
}
