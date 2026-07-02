/** Sanitized PDF filename for a single RFI export. */
export function rfiPdfFileName(rfiNum: string): string {
  return (rfiNum || 'RFI').replace(/[^a-zA-Z0-9_-]/g, '_') + '.pdf';
}
