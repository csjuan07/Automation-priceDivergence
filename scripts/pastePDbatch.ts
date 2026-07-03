const MS_PER_DAY = 24 * 60 * 60 * 1000;
const OUTPUT_SCHEMA_SIZE = 29;
const CHUNK_SIZE = 400;

function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  partNumbers: string[],
  supplierIds: string[],
  supplierNames: string[],
  prices: string[],
  currencies: string[],
  volumes: string[],
  descriptions: string[],
  Category: string[],
  invoiceDate: string
): string {

  const table = workbook.getTable(tableName);

  if (!table) {
    throw new Error(`Tabela '${tableName}' não encontrada.`);
  }

  const importDate = parseBRDateToDate(invoiceDate);

  if (isNaN(importDate.getTime())) {
    throw new Error("Data inválida. Utilize o formato dd/mm/aaaa.");
  }

  const arrPartNumbers = Array.isArray(partNumbers) ? partNumbers : [];
  const arrSupplierIds = Array.isArray(supplierIds) ? supplierIds : [];
  const arrSupplierNames = Array.isArray(supplierNames) ? supplierNames : [];
  const arrPrices = Array.isArray(prices) ? prices : [];
  const arrCurrencies = Array.isArray(currencies) ? currencies : [];
  const arrVolumes = Array.isArray(volumes) ? volumes : [];
  const arrDescriptions = Array.isArray(descriptions) ? descriptions : [];
  const arrCategory = Array.isArray(Category) ? Category : [];

  const rowCount = arrPartNumbers.length;

  if (
    rowCount === 0 ||
    arrSupplierIds.length !== rowCount ||
    arrSupplierNames.length !== rowCount ||
    arrPrices.length !== rowCount ||
    arrCurrencies.length !== rowCount ||
    arrVolumes.length !== rowCount ||
    arrDescriptions.length !== rowCount ||
    arrCategory.length !== rowCount
  ) {
    throw new Error(
      `Tamanhos diferentes entre os arrays. ` +
      `partNumbers=${arrPartNumbers.length}, ` +
      `supplierIds=${arrSupplierIds.length}, ` +
      `supplierNames=${arrSupplierNames.length}, ` +
      `prices=${arrPrices.length}, ` +
      `currencies=${arrCurrencies.length}, ` +
      `volumes=${arrVolumes.length}, ` +
      `descriptions=${arrDescriptions.length}, ` +
      `Category=${arrCategory.length}`
    );
  }

  const tableColumnCount = table.getColumns().length;

  if (tableColumnCount !== OUTPUT_SCHEMA_SIZE) {
    throw new Error(
      `A tabela possui ${tableColumnCount} colunas, ` +
      `mas o processo espera ${OUTPUT_SCHEMA_SIZE}.`
    );
  }

  const rows: (string | number | boolean | null)[][] =
    new Array(rowCount);

  const excelDate =
    dateToExcelSerial(importDate);

  for (let i = 0; i < rowCount; i++) {

    rows[i] = [
      toStringOrEmpty(arrPartNumbers[i]),
      toNumberOrNull(arrPrices[i]),
      toStringOrEmpty(arrCurrencies[i]),
      null,
      null,
      excelDate,
      toNumberOrNull(arrVolumes[i]),
      null,
      null,
      toNumberOrNull(arrSupplierIds[i]),
      toStringOrEmpty(arrSupplierNames[i]),
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      toStringOrEmpty(arrDescriptions[i]),
      null,
      null,
      null,
      toStringOrEmpty(arrCategory[i])
    ];
  }

  const hadTotals = table.getShowTotals();

  if (hadTotals) {
    table.setShowTotals(false);
  }

  let batches = 0;

  try {

    for (
      let start = 0;
      start < rows.length;
      start += CHUNK_SIZE
    ) {

      const end =
        Math.min(start + CHUNK_SIZE, rows.length);

      const batch =
        rows.slice(start, end);

      table.addRows(-1, batch);

      batches++;
    }

  } finally {

    if (hadTotals) {
      table.setShowTotals(true);
    }

  }

  return (
    `Importação concluída com sucesso: ` +
    `${rowCount} linhas processadas em ${batches} lote(s).`
  );
}

function toNumberOrNull(
  value: string | undefined
): number | null {

  if (
    value === undefined ||
    value === null ||
    value === ""
  ) {
    return null;
  }

  const numberValue = Number(value);

  return Number.isFinite(numberValue)
    ? numberValue
    : null;
}

function toStringOrEmpty(
  value: string | undefined
): string {

  return value ?? "";
}

function dateToExcelSerial(
  date: Date
): number {

  const utc = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );

  const base =
    Date.UTC(1899, 11, 30);

  return (
    utc - base
  ) / MS_PER_DAY;
}

function parseBRDateToDate(
  value: string
): Date {

  const parts =
    value.split("/");

  if (parts.length !== 3) {
    return new Date("invalid");
  }

  const day =
    Number(parts[0]);

  const month =
    Number(parts[1]);

  const year =
    Number(parts[2]);

  return new Date(
    year,
    month - 1,
    day
  );
}