import { seedData } from "./mockData";
import type { AppData } from "./types";

export const STORAGE_KEY = "circlia-step1-v1";

function normalizeData(data: AppData): AppData {
  data.cases = data.cases.map((item) => {
    const initiator = data.approvers.find((person) => person.id === item.initiatorId);
    const members = item.members.some((member) => member.isInitiator) ? item.members : [
      {
        id: `case-member-initiator-${item.id}`,
        approverId: item.initiatorId,
        name: item.initiatorName,
        email: initiator?.email ?? "",
        department: initiator?.department ?? "",
        sequence: 1,
        requiresStamp: false,
        isInitiator: true,
        status: "approved" as const,
        completedAt: item.startedAt,
      },
      ...item.members.map((member, index) => ({ ...member, sequence: index + 2 })),
    ];
    return {
      ...item,
      members,
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
    };
  });
  return data;
}

export function loadData(): AppData {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    const data = saved ? JSON.parse(saved) as AppData : structuredClone(seedData);
    return normalizeData(data);
  } catch {
    return normalizeData(structuredClone(seedData));
  }
}

export function saveData(data: AppData) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

export function resetData(): AppData {
  const fresh = normalizeData(structuredClone(seedData));
  saveData(fresh);
  return fresh;
}
