function main(
  workbook: ExcelScript.Workbook,
  sheetName: string,
  columnName: string,
  startRow: number,
  lastRow: string
): string[] {

  const worksheet =
    workbook.getWorksheet(sheetName);

  if (!worksheet) {
    throw new Error(
      `Worksheet '${sheetName}' not found.`
    );
  }

  const rowLimit = Number(lastRow);

  if (!Number.isFinite(rowLimit)) {
    throw new Error(
      "Invalid lastRow value."
    );
  }

  const lastIndex =
    rowLimit - 1;

  const range =
    worksheet.getRange(
      `${columnName}${startRow}:${columnName}${lastIndex}`
    );

  const values =
    range.getValues();

  return columnToArray(values);
}

function columnToArray(
  values: (string | number | boolean | null)[][]
): string[] {

  const result: string[] = [];

  for (
    let i = 0;
    i < values.length;
    i++
  ) {
    result.push(
      String(values[i][0] ?? "")
    );
  }

  return result;
}