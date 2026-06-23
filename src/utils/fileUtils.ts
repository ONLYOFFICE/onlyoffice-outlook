import type { DocFormat } from "../types";

export class FileUtils {
  private formats: DocFormat[];

  constructor(formats: DocFormat[]) {
    this.formats = formats;
  }

  getFormats(): DocFormat[] {
    return this.formats;
  }

  getFileExtension(fileName: string): string {
    if (!fileName || !fileName.includes(".")) return "";
    return fileName.split(".").pop()!.toLowerCase();
  }
}
