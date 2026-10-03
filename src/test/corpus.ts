/** Office owner/lock files are not documents, even though their extensions match. */
export function isOfficeCorpusFile(name: string): boolean {
  return !name.startsWith('~$') && /\.(docx|xlsx|pptx)$/i.test(name)
}
