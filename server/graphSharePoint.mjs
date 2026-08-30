const DEFAULT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const SUPPORTED_FILE_TYPES = new Map([
  ["pdf", "PDF"],
  ["xls", "Excel"],
  ["xlsx", "Excel"],
  ["xlsm", "Excel"],
  ["doc", "Word"],
  ["docx", "Word"],
  ["ppt", "PowerPoint"],
  ["pptx", "PowerPoint"],
]);

function encodePath(value) {
  return String(value || "")
    .split("/")
    .filter(Boolean)
    .map((part) => encodeURIComponent(part))
    .join("/");
}

function normalizeSitePath(value) {
  const normalized = String(value || "").trim().replace(/\\/g, "/");
  return `/${normalized.split("/").filter(Boolean).join("/")}`;
}

function fileType(name) {
  const extension = String(name || "").split(".").at(-1)?.toLowerCase() || "";
  return SUPPORTED_FILE_TYPES.get(extension);
}

function priority(type) {
  return type === "PDF" ? 0 : type === "Excel" ? 1 : 2;
}

function normalizeFile(item, location) {
  const type = fileType(item.name);
  if (!item.id || !item.file || !type || !item.webUrl) return undefined;
  return {
    id: String(item.id),
    provider: "sharepoint",
    name: String(item.name || "ファイル名未設定"),
    type,
    location,
    updatedAt: String(item.lastModifiedDateTime || ""),
    updatedBy: String(item.lastModifiedBy?.user?.displayName || item.lastModifiedBy?.application?.displayName || "更新者不明"),
    fileUrl: String(item.webUrl),
  };
}

export function isSharePointTargetConfigured(env = process.env) {
  const hostname = env.SHAREPOINT_HOSTNAME?.trim() || env.VITE_SHAREPOINT_HOSTNAME?.trim();
  const sitePath = env.SHAREPOINT_SITE_PATH?.trim() || env.VITE_SHAREPOINT_SITE_PATH?.trim();
  return Boolean(hostname && sitePath);
}

export function createGraphSharePoint({
  acquireAccessToken,
  hostname,
  sitePath,
  libraryName = "",
  folderPath = "",
  fetchImpl = fetch,
  graphBaseUrl = DEFAULT_GRAPH_BASE_URL,
  cacheTtlMs = 60_000,
  maxFiles = 500,
  now = Date.now,
} = {}) {
  if (typeof acquireAccessToken !== "function") throw new Error("SharePoint access token provider is required");
  if (!String(hostname || "").trim() || !String(sitePath || "").trim()) throw new Error("SharePoint target site is not configured");

  const normalizedHostname = String(hostname).trim();
  const normalizedSitePath = normalizeSitePath(sitePath);
  const normalizedLibraryName = String(libraryName || "").trim();
  const normalizedFolderPath = String(folderPath || "").trim().replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
  const normalizedCacheTtlMs = Number.isFinite(cacheTtlMs) && cacheTtlMs >= 0 ? cacheTtlMs : 60_000;
  const normalizedMaxFiles = Number.isFinite(maxFiles) && maxFiles > 0 ? Math.floor(maxFiles) : 500;
  let cachedResult;
  let expiresAt = 0;
  let inFlight;

  async function requestJson(url, token) {
    const response = await fetchImpl(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!response.ok) {
      const requestId = response.headers?.get?.("request-id") || response.headers?.get?.("client-request-id");
      throw new Error(`Microsoft Graph SharePoint request failed (${response.status})${requestId ? ` request-id=${requestId}` : ""}`);
    }
    return response.json();
  }

  async function loadFiles() {
    if (cachedResult && now() < expiresAt) return cachedResult;
    if (inFlight) return inFlight;

    inFlight = (async () => {
      const token = await acquireAccessToken();
      const siteUrl = `${graphBaseUrl}/sites/${encodeURIComponent(normalizedHostname)}:${normalizedSitePath.split("/").map(encodeURIComponent).join("/")}?$select=id,displayName,webUrl`;
      const site = await requestJson(siteUrl, token);
      if (!site.id) throw new Error("SharePoint site was not returned");

      const drivesUrl = `${graphBaseUrl}/sites/${encodeURIComponent(site.id)}/drives?$select=id,name,webUrl,driveType&$top=200`;
      const drivesPayload = await requestJson(drivesUrl, token);
      const drives = Array.isArray(drivesPayload.value) ? drivesPayload.value : [];
      const library = normalizedLibraryName
        ? drives.find((drive) => String(drive.name || "").localeCompare(normalizedLibraryName, "ja", { sensitivity: "accent" }) === 0)
        : drives.find((drive) => drive.driveType === "documentLibrary") || drives[0];
      if (!library?.id) throw new Error(normalizedLibraryName ? `SharePoint document library was not found: ${normalizedLibraryName}` : "SharePoint document library was not found");

      const encodedFolder = encodePath(normalizedFolderPath);
      const childrenPath = encodedFolder ? `root:/${encodedFolder}:/children` : "root/children";
      const select = "id,name,webUrl,file,size,lastModifiedDateTime,lastModifiedBy";
      let nextUrl = `${graphBaseUrl}/drives/${encodeURIComponent(library.id)}/${childrenPath}?$select=${encodeURIComponent(select)}&$top=200`;
      const files = [];
      const location = [site.displayName || normalizedSitePath, library.name, normalizedFolderPath].filter(Boolean).join(" / ");

      while (nextUrl && files.length < normalizedMaxFiles) {
        const payload = await requestJson(nextUrl, token);
        for (const item of payload.value || []) {
          const file = normalizeFile(item, location);
          if (file) files.push(file);
          if (files.length >= normalizedMaxFiles) break;
        }
        nextUrl = typeof payload["@odata.nextLink"] === "string" ? payload["@odata.nextLink"] : "";
      }

      files.sort((left, right) => priority(left.type) - priority(right.type) || left.name.localeCompare(right.name, "ja"));
      cachedResult = {
        siteName: String(site.displayName || normalizedSitePath),
        siteUrl: String(site.webUrl || ""),
        libraryName: String(library.name || "ドキュメント"),
        folderPath: normalizedFolderPath,
        files,
      };
      expiresAt = now() + normalizedCacheTtlMs;
      return cachedResult;
    })().finally(() => {
      inFlight = undefined;
    });

    return inFlight;
  }

  return {
    listFiles: loadFiles,
    clearCache() {
      cachedResult = undefined;
      expiresAt = 0;
    },
  };
}

export { normalizeFile };
