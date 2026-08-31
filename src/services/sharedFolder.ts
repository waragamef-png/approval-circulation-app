const configuredGateway = (import.meta.env.VITE_SHARED_FOLDER_API_BASE_URL || "").trim();
const configuredApi = (import.meta.env.VITE_API_BASE_URL || "").trim();
const apiBaseUrl = (configuredGateway || configuredApi) === "same-origin"
  ? ""
  : (configuredGateway || configuredApi).replace(/\/$/, "");

const supportedFilePattern = /\.(pdf|xlsx?|xlsm|docx?|pptx?)$/i;

export function normalizeSharedFolderPath(value: string) {
  const candidate = value.trim().replaceAll("/", "\\");
  const isUncPath = /^\\\\[^\\]+\\[^\\]+\\.+/.test(candidate);
  const isDrivePath = /^[a-z]:\\.+/i.test(candidate);
  if ((!isUncPath && !isDrivePath) || !supportedFilePattern.test(candidate)) return undefined;
  return candidate;
}

export function fileNameFromSharedFolderPath(value: string) {
  const normalized = normalizeSharedFolderPath(value);
  return normalized?.split("\\").filter(Boolean).at(-1) || "";
}

export function sharedFolderOpenUrl(filePath: string) {
  const normalized = normalizeSharedFolderPath(filePath);
  if (!normalized) return undefined;
  return `${apiBaseUrl}/api/shared-files/open?path=${encodeURIComponent(normalized)}`;
}
