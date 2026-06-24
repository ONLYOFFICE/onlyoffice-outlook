import sign from "jwt-encode";
import type { DocFormat } from "../types";

const DOCUMENT_SERVER_JWT_EXPIRES_IN = 60;

export class FileUtils {
  private formats: DocFormat[];

  constructor(formats: DocFormat[]) {
    this.formats = formats;
  }

  getFormats(): DocFormat[] {
    return this.formats;
  }

  getDocFormatByExtension(
    extension: string,
  ): DocFormat | undefined {
    return this.formats.find((f) => extension === f.name);
  }

  getExtension(fileName: string): string {
    if (!fileName || !fileName.includes(".")) return "";
    return fileName.split(".").pop()!.toLowerCase();
  }

  getDocumentType(
    extension: string,
  ): string | null {
    const format = this.getDocFormatByExtension(extension);

    if (format?.type) return format.type;

    return null;
  }

  isViewable(extension: string): boolean {
    return Boolean(this.getDocumentType(extension));
  }

  isEditable(extension: string): boolean {
    const format = this.getDocFormatByExtension(extension);

    return format?.actions.includes("edit") ?? false;
  }

  async createKey(value: string) {
    const buffer = await crypto.subtle.digest(
      'SHA-256',
      new TextEncoder().encode(value)
    );

    return Array.from(new Uint8Array(buffer))
      .map(b => b.toString(16).padStart(2, '0'))
      .join('');
  }

  createEditorConfig(
    key: string,
    fileName: string,
    fileUrl: string,
    mode: "edit" | "view",
    user: {
      id: string,
      name: string,
    },
    secret: string,
    locale: string = "en",
  ) {
    const extension = this.getExtension(fileName);

    const config = {
      document: {
        fileType: extension,
        key,
        title: fileName,
        url: fileUrl,
        permissions: {
          edit: this.isEditable(extension),
        },
      },
      documentType: this.getDocumentType(extension),
      editorConfig: {
        mode,
        customization: {
          forcesave: true,
        },
        user,
        lang: locale,
      },
      token: null as string | null,
    };

    if (secret) {
      const exp = Math.floor(Date.now() / 1000) + DOCUMENT_SERVER_JWT_EXPIRES_IN;
      config.token = sign({ ...config, exp }, secret);
    }

    return config;
  }
}
