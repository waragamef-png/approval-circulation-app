import { seedData } from "./mockData";
import type { AppData } from "./types";

const STORAGE_KEY = "circlia-step1-v1";

export function loadData(): AppData {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return structuredClone(seedData);
    const parsed = JSON.parse(saved) as AppData;
    parsed.cases = parsed.cases.map((item) => ({
      ...item,
      accessKey: item.accessKey || crypto.randomUUID().replaceAll("-", ""),
    }));
    return parsed;
  } catch {
    return structuredClone(seedData);
  }
}

export function saveData(data: AppData) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

export function resetData(): AppData {
  const fresh = structuredClone(seedData);
  saveData(fresh);
  return fresh;
}
