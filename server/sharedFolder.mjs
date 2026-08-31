import { realpath, stat } from "node:fs/promises";
import path from "node:path";

const supportedExtensions = new Set([".pdf", ".xls", ".xlsx", ".xlsm", ".doc", ".docx", ".ppt", ".pptx"]);

export class SharedFolderError extends Error {
  constructor(message, statusCode) {
    super(message);
    this.name = "SharedFolderError";
    this.statusCode = statusCode;
  }
}

export function parseSharedFolderRoots(value = "") {
  return value
    .split(";")
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => path.resolve(item));
}

function isWithinRoot(rootPath, filePath) {
  const relative = path.relative(rootPath, filePath);
  return relative === "" || (!relative.startsWith(`..${path.sep}`) && relative !== ".." && !path.isAbsolute(relative));
}

export function createSharedFolderGateway({ roots = [], fileSystem = { realpath, stat } } = {}) {
  const configuredRoots = roots.map((item) => path.resolve(item));

  return {
    configured: configuredRoots.length > 0,

    async resolveFile(requestedPath) {
      if (configuredRoots.length === 0) {
        throw new SharedFolderError("共有フォルダが設定されていません。", 503);
      }

      const candidate = String(requestedPath || "").trim();
      if (!candidate || !path.isAbsolute(candidate)) {
        throw new SharedFolderError("共有フォルダの完全なファイルパスを指定してください。", 400);
      }

      if (!supportedExtensions.has(path.extname(candidate).toLowerCase())) {
        throw new SharedFolderError("PDF、Excel、Word、PowerPointのファイルを指定してください。", 400);
      }

      const resolvedCandidate = path.resolve(candidate);
      if (!configuredRoots.some((root) => isWithinRoot(root, resolvedCandidate))) {
        throw new SharedFolderError("このファイルは許可された共有フォルダ内にありません。", 403);
      }

      let realFilePath;
      let fileStat;
      try {
        realFilePath = await fileSystem.realpath(resolvedCandidate);
        fileStat = await fileSystem.stat(realFilePath);
      } catch (error) {
        if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
          throw new SharedFolderError("共有フォルダにファイルが見つかりません。", 404);
        }
        throw error;
      }

      if (!fileStat.isFile()) {
        throw new SharedFolderError("ファイルを指定してください。", 400);
      }

      const realRoots = [];
      for (const root of configuredRoots) {
        try {
          realRoots.push(await fileSystem.realpath(root));
        } catch {
          // A temporarily unavailable share must not make another configured root unusable.
        }
      }

      if (!realRoots.some((root) => isWithinRoot(root, realFilePath))) {
        throw new SharedFolderError("このファイルは許可された共有フォルダ内にありません。", 403);
      }

      return {
        absolutePath: realFilePath,
        fileName: path.basename(realFilePath),
        inline: path.extname(realFilePath).toLowerCase() === ".pdf",
      };
    },
  };
}

export function contentDisposition(fileName, inline) {
  const fallback = fileName.replace(/[^\x20-\x7e]/g, "_").replace(/["\\]/g, "_");
  const encoded = encodeURIComponent(fileName).replace(/[!'()*]/g, (character) => `%${character.charCodeAt(0).toString(16).toUpperCase()}`);
  return `${inline ? "inline" : "attachment"}; filename="${fallback}"; filename*=UTF-8''${encoded}`;
}
