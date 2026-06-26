import pdfBgImg from '../../webparts/managerDashboard/assets/pdf-backgroundimage.png';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function drawPdfBg(doc: any, pw: number, ph: number): void {
  doc.addImage(pdfBgImg, 'PNG', 0, 0, pw, ph);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function drawLetterhead(doc: any, pw: number, ph: number, title: string, subtitle: string): number {
  drawPdfBg(doc, pw, ph);
  const barY = 40;
  const barH = 12;
  doc.setFillColor(26, 32, 48);
  doc.rect(0, barY, pw, barH, 'F');
  doc.setFillColor(42, 158, 42);
  doc.rect(0, barY, 3, barH, 'F');
  doc.setFontSize(11);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(255, 255, 255);
  doc.text(title, 8, barY + 7.5);
  if (subtitle) {
    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(160, 175, 195);
    doc.text(subtitle, pw - 8, barY + 7.5, { align: 'right' });
  }
  return barY + barH + 4;
}
