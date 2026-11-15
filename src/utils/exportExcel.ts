import * as XLSX from 'xlsx';

export interface ExcelSheetData {
  sheetName: string;
  headers: string[];
  rows: (string | number)[][];
  footerRows?: (string | number)[][];
}

export function appendRtlSheet(
  wb: XLSX.WorkBook,
  sheet: ExcelSheetData
): void {
  const allRows = [
    sheet.headers,
    ...sheet.rows,
    ...(sheet.footerRows ?? []),
  ];
  const ws = XLSX.utils.aoa_to_sheet(allRows);
  ws['!views'] = [{ RTL: true }];
  ws['!cols'] = sheet.headers.map((h, i) => {
    const maxLen = Math.max(
      h.length,
      ...sheet.rows.map((r) => String(r[i] ?? '').length),
      ...(sheet.footerRows ?? []).map((r) => String(r[i] ?? '').length)
    );
    return { wch: Math.min(Math.max(maxLen + 5, 13), 42) };
  });
  XLSX.utils.book_append_sheet(wb, ws, sheet.sheetName.slice(0, 31));
}

export function exportToExcel(fileName: string, sheets: ExcelSheetData[]): void {
  const wb = XLSX.utils.book_new();
  wb.Workbook = { Views: [{ RTL: true }] };

  for (const sheet of sheets) {
    appendRtlSheet(wb, sheet);
  }

  XLSX.writeFile(wb, `${fileName}.xlsx`);
}

