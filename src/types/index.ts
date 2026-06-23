export type DocFormatType = "word" | "cell" | "slide" | "pdf" | "diagram" | "";

export interface DocFormat {
  name: string;
  type: DocFormatType;
  actions: string[];
  convert: string[];
  mime: string[];
}
