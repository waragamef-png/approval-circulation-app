export interface DirectoryUser {
  id: string;
  name: string;
  email: string;
  department: string;
}

interface DirectoryResponse {
  users: DirectoryUser[];
}

const configuredValue = (import.meta.env.VITE_API_BASE_URL || "").trim();
const configuredBaseUrl = configuredValue === "same-origin" ? "" : configuredValue.replace(/\/$/, "");

export const directorySearchEnabled = configuredValue.length > 0;

export async function searchDirectoryUsers(query: string, signal?: AbortSignal): Promise<DirectoryUser[]> {
  const normalized = query.trim();
  if (!directorySearchEnabled || !normalized) return [];
  const response = await fetch(`${configuredBaseUrl}/api/directory/users?q=${encodeURIComponent(normalized)}&limit=10`, { signal });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({})) as { message?: string };
    throw new Error(payload.message || "Microsoft 365の社内名簿を検索できませんでした。");
  }
  const payload = await response.json() as DirectoryResponse;
  return payload.users;
}
