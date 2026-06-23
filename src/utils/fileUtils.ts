export function getFileExtension(fileName: string): string {
  if (!fileName || !fileName.includes(".")) return "";
  return fileName.split(".").pop()!.toLowerCase();
}
