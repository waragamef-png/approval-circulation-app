import type { MockFile, ProviderType } from "../types";

export interface StorageProvider {
  readonly type: ProviderType;
  listFiles(): Promise<MockFile[]>;
  getFile(fileId: string): Promise<MockFile | undefined>;
  openFile(fileId: string): Promise<string>;
  downloadFile(fileId: string): Promise<Blob>;
  updateFile(fileId: string, content: Blob): Promise<void>;
  getMetadata(fileId: string): Promise<MockFile | undefined>;
  getWebUrl(fileId: string): Promise<string>;
}

// STEP 1ではUIのみが対象です。後続STEPでGraph/社内API実装をこの契約へ接続します。
