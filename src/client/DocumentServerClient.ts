import type { DocFormat } from "../types";
import formats from "../../assets/onlyoffice-docs-formats.json";

export class DocumentServerClient {
  async getFormats(): Promise<DocFormat[]> {
    return formats as DocFormat[];
  }
}
