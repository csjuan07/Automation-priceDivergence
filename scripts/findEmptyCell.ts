function main(
  workbook: ExcelScript.Workbook,
  sheetName: string,
  columnLetter: string,
  startRow: number
): string {

  const worksheet =
    workbook.getWorksheet(sheetName);

  if (!worksheet) {
    throw new Error(
      `Worksheet '${sheetName}' not found.`
    );
  }

  const usedRange =
    worksheet.getUsedRange();

  let lastRow = startRow;

  if (usedRange) {
    lastRow =
      usedRange.getRowCount();
  }

  const range =
    worksheet.getRange(
      `${columnLetter}${startRow}:${columnLetter}${lastRow}`
    );

  const values =
    range.getValues();

  for (
    let rowIndex = 0;
    rowIndex < values.length;
    rowIndex++
  ) {

    const value =
      values[rowIndex][0];

    const isEmpty =
      value === null ||
      (
        typeof value === "string" &&
        value.trim() === ""
      );

    if (isEmpty) {

      const worksheetRow =
        startRow + rowIndex;

      return String(
        worksheetRow
      );
    }
  }

  return String(lastRow);
}