import html2canvas from 'html2canvas';

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
}

function toCsvCell(v: unknown): string {
  const s = String(v ?? '');
  const needsQuote = /[\n\r\t,\"]/g.test(s);
  const escaped = s.replace(/\"/g, '""');
  return needsQuote ? `"${escaped}"` : escaped;
}

export function exportCsv(
  filenameBase: string,
  headers: string[],
  rows: Array<Array<unknown>>
) {
  const lines: string[] = [];
  lines.push(headers.map(toCsvCell).join(','));
  for (const r of rows) {
    lines.push(r.map(toCsvCell).join(','));
  }
  const csv = lines.join('\n');
  const blob = new Blob(["\uFEFF", csv], { type: 'text/csv;charset=utf-8' });
  downloadBlob(blob, `${filenameBase}.csv`);
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function exportExcelHtml(
  filenameBase: string,
  headers: string[],
  rows: Array<Array<unknown>>
) {
  const thead = `<tr>${headers.map((h) => `<th>${escapeHtml(String(h ?? ''))}</th>`).join('')}</tr>`;
  const tbody = rows
    .map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(String(c ?? ''))}</td>`).join('')}</tr>`)
    .join('');

  const html = `<!doctype html><html><head><meta charset="utf-8" /></head><body><table>${thead}${tbody}</table></body></html>`;
  const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8' });
  downloadBlob(blob, `${filenameBase}.xls`);
}

export async function exportElementPng(element: HTMLElement, filenameBase: string) {
  // 暫時移除滾動限制，讓 html2canvas 捕捉完整內容
  const originalOverflow = element.style.overflow;
  const originalMaxHeight = element.style.maxHeight;
  const originalHeight = element.style.height;
  element.style.overflow = 'visible';
  element.style.maxHeight = 'none';
  element.style.height = 'auto';

  try {
    const canvas = await html2canvas(element, {
      backgroundColor: '#FFFFFF',
      scale: 2,
      scrollX: 0,
      scrollY: 0,
      windowWidth: element.scrollWidth,
      windowHeight: element.scrollHeight,
    });
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (!blob) return;
    downloadBlob(blob, `${filenameBase}.png`);
  } finally {
    // 還原原本的樣式
    element.style.overflow = originalOverflow;
    element.style.maxHeight = originalMaxHeight;
    element.style.height = originalHeight;
  }
}
