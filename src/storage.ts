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
      documents: item.documents?.length ? item.documents : [{
        id: item.fileId,
        fileId: item.fileId,
        name: item.fileName,
        type: item.fileName.toLowerCase().endsWith(".pdf") ? "PDF" : "Excel",
        location: "保存場所未設定",
        fileUrl: item.fileUrl,
        requiresStamp: item.members.some((member) => member.requiresStamp),
      }],
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
