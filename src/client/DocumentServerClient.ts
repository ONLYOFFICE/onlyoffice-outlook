/* global fetch */

import type { DocFormat } from "../types";

export class DocumentServerClient {
  private readonly baseUrl: string;

  constructor(baseUrl: string) {
    this.baseUrl = baseUrl.replace(/\/$/, "");
  }

  async getFormats(): Promise<DocFormat[]> {
    const response = await fetch(`${this.baseUrl}/meta/formats`);

    if (!response.ok) {
      throw new Error(`Failed to fetch formats: ${response.status} ${response.statusText}`);
    }

    return response.json() as Promise<DocFormat[]>;
  }
}
