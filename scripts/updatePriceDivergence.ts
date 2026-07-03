const MS_PER_DAY = 24 * 60 * 60 * 1000;

// ========= Tipos auxiliares =========
type Cell = string | number | boolean | null;
type Matrix = Cell[][];
type NumMatrix = (number | null)[][];
type StrMatrix = (string | null)[][];
type ERPItem = { article: string; supplierId: number; price: number; currency: string; measure: string };

// ========= Utilitários puros e tipados =========

function toStringSafe(v: unknown): string {
    if (v === null || v === undefined) return "";
    return String(v);
}

function toUpperCode(v: unknown): string {
    return toStringSafe(v).trim().toUpperCase();
}

function toNumberSafe(v: unknown): number | null {
    if (typeof v === "number") {
        return Number.isFinite(v) ? v : null;
    }
    if (typeof v === "boolean") {
        return v ? 1 : 0;
    }
    const s = toStringSafe(v).trim();
    if (!s) return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
}

function parseBRDateToDate(dataBR: string): Date {
    const onlyDate = dataBR.split(" ")[0]?.trim() ?? dataBR;
    const parts = onlyDate.split("/");
    if (parts.length !== 3) {
        // fallback: tenta Date.parse (não recomendado, mas evita crash)
        const d = new Date(onlyDate);
        return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
    const [ddStr, mmStr, yyyyStr] = parts;
    const dd = Number(ddStr);
    const mm = Number(mmStr);
    const yyyy = Number(yyyyStr);
    return new Date(yyyy, mm - 1, dd);
}

function dateToExcelSerial(d: Date): number {
    // Zera hora em UTC para evitar fuso
    const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
    const base = Date.UTC(1899, 11, 30); // 1899-12-30
    return (utc - base) / MS_PER_DAY;
}

function getSerialFromCellValue(v: unknown): number | null {
    // Aceita já como número (serial) ou como string dd/mm/aaaa
    const num = toNumberSafe(v);
    if (num !== null) return num;

    const s = toStringSafe(v).trim();
    if (!s) return null;
    const d = parseBRDateToDate(s);
    if (isNaN(d.getTime())) return null;
    return dateToExcelSerial(d);
}

function cloneValues(rng: ExcelScript.Range): Matrix {
    // Copia a matriz para permitir edições antes de setValues
    const values = rng.getValues(); // retorna (string|number|boolean)[][]
    const out: Matrix = new Array(values.length);
    for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const clonedRow: Cell[] = new Array(row.length);
        for (let j = 0; j < row.length; j++) clonedRow[j] = row[j] as Cell;
        out[i] = clonedRow;
    }
    return out;
}

function main(workbook: ExcelScript.Workbook, mesAtual: string): string {

    let countZero: number;
    let countTotal: number;

    let maiorDiffRel = 0;
    let maiorDiffAbs = 0;
    // ========= Parâmetro e serial do mês =========
    const mesSerial = dateToExcelSerial(parseBRDateToDate(mesAtual));

    // ========= Tabelas necessárias =========
    const tableRates = workbook.getTable("ExchangeRates");
    const tablePD = workbook.getTable("PriceDivergence");
    const tableERP = workbook.getTable("ERPReference");

    const tableRefs = workbook.getTable("ReferenceData");
    if (!tableRefs) {
        throw new Error("Tabela 'ReferenceData' não encontrada.");
    }


    if (!tableRates || !tablePD || !tableERP) {
        throw new Error("Uma ou mais tabelas não foram encontradas (ExchanceRates, PriceDivergence, ERPReference). Verifique os nomes.");
    }

    // (Opcional) Acelerar desligando Totals Row
    const hadTotalsPD = tablePD.getShowTotals();
    if (hadTotalsPD) tablePD.setShowTotals(false);

    // ========= EXCHANGE RATES =========
    const rngCurrency = tableRates.getColumnByName("CurrencyCode").getRangeBetweenHeaderAndTotal();
    const rngRates = tableRates.getColumnByName("Rates2026").getRangeBetweenHeaderAndTotal();
    const curVals = rngCurrency.getValues(); // Matrix
    const rateVals = rngRates.getValues();   // Matrix

    const ratesMap = new Map<string, number>();
    for (let i = 0; i < curVals.length; i++) {
        const cur = toUpperCode(curVals[i][0] as unknown);
        const rate = toNumberSafe(rateVals[i][0] as unknown);
        if (cur && rate !== null) ratesMap.set(cur, rate);
    }

    const colRefArticle = tableRefs.getColumnByName("Article").getRangeBetweenHeaderAndTotal();
    const colRefArticle_SC = tableRefs.getColumnByName("Article S/ C").getRangeBetweenHeaderAndTotal();
    const colRefRG = tableRefs.getColumnByName("REGIONAL / GLOBAL").getRangeBetweenHeaderAndTotal();
    const colRefID = tableRefs.getColumnByName("ID").getRangeBetweenHeaderAndTotal();
    const colRefCM = tableRefs.getColumnByName("CM").getRangeBetweenHeaderAndTotal();

    const refIDvals = colRefID.getValues();
    const refCMvals = colRefCM.getValues();
    const refArticlevals = colRefArticle_SC.getValues();
    const refRGvals = colRefRG.getValues();

    // ========= NET PRICES (ERP) =========
    const ERPArticlerng = tableERP.getColumnByName("Article ERP").getRangeBetweenHeaderAndTotal();
    const ERPIDrng = tableERP.getColumnByName("ID ERP").getRangeBetweenHeaderAndTotal();
    const ERPPricerng = tableERP.getColumnByName("PRICE ERP").getRangeBetweenHeaderAndTotal();
    const ERPCurrng = tableERP.getColumnByName("MOEDA ERP").getRangeBetweenHeaderAndTotal();
    const ERPUMrng = tableERP.getColumnByName("UM ERP").getRangeBetweenHeaderAndTotal();

    const ERPArticlevals = ERPArticlerng.getValues();
    const ERPIDvals = ERPIDrng.getValues();
    const ERPPricevals = ERPPricerng.getValues();
    const ERPCurvals = ERPCurrng.getValues();
    const ERPUMvals = ERPUMrng.getValues();

    const ERPByArticleSup = new Map<string, ERPItem>(); // "Article|ID" -> item
    const ERPByArticle = new Map<string, ERPItem>();    // "Article" -> item

    const ERPLen = ERPArticlevals.length;
    for (let i = 0; i < ERPLen; i++) {
        const article = toUpperCode(ERPArticlevals[i][0] as unknown);
        if (!article) continue;
        const supplierId = toNumberSafe(ERPIDvals[i][0] as unknown);
        const price = toNumberSafe(ERPPricevals[i][0] as unknown);
        const currency = toUpperCode(ERPCurvals[i][0] as unknown);
        const measure = toStringSafe(ERPUMvals[i][0] as unknown);

        const item: ERPItem = {
            article,
            supplierId: supplierId ?? NaN,
            price: price ?? NaN,
            currency,
            measure
        };
        ERPByArticle.set(article, item);
        if (supplierId !== null && Number.isFinite(supplierId)) {
            ERPByArticleSup.set(`${article}|${supplierId}`, item);
        }
    }

    // ========= Colunas da PD =========
    const colArticle = tablePD.getColumnByName("Article - ACTUAL").getRangeBetweenHeaderAndTotal();
    const colArticleSC = tablePD.getColumnByName("Article S/ C").getRangeBetweenHeaderAndTotal();
    const colSupID = tablePD.getColumnByName("SUPPLIER ID ACTUAL").getRangeBetweenHeaderAndTotal();
    const colPriceCtrl = tablePD.getColumnByName("PRICE - ACTUAL").getRangeBetweenHeaderAndTotal();
    const colCurCtrl = tablePD.getColumnByName("CURRENCY - ACTUAL").getRangeBetweenHeaderAndTotal();
    const colSekCtrl = tablePD.getColumnByName("ACTUAL PRICE IN SEK").getRangeBetweenHeaderAndTotal();
    const colMeasureActual = tablePD.getColumnByName("MEASURE - ACTUAL").getRangeBetweenHeaderAndTotal();
    const colVol = tablePD.getColumnByName("VOLUME").getRangeBetweenHeaderAndTotal();
    const colSpendAct = tablePD.getColumnByName("SPEND - ACTUAL").getRangeBetweenHeaderAndTotal();
    const colSpendActSek = tablePD.getColumnByName("SPEND - ACTUAL SEK").getRangeBetweenHeaderAndTotal();
    const colBuyer = tablePD.getColumnByName("BUYER").getRangeBetweenHeaderAndTotal();

    const colDiffAbs = tablePD.getColumnByName("|DIFERENÇA SEK|").getRangeBetweenHeaderAndTotal();
    const colDiffRel = tablePD.getColumnByName("|DIFERENÇA%|").getRangeBetweenHeaderAndTotal();

    const colPriceERP = tablePD.getColumnByName("PRICE - ERP").getRangeBetweenHeaderAndTotal();
    const colCurERP = tablePD.getColumnByName("CURRENCY - ERP").getRangeBetweenHeaderAndTotal();
    const colPricePCERP = tablePD.getColumnByName("PRICE - ERP PC").getRangeBetweenHeaderAndTotal();
    const colSekERP = tablePD.getColumnByName("ERP PRICE IN SEK").getRangeBetweenHeaderAndTotal();
    const colSpendERP = tablePD.getColumnByName("SPEND - ERP").getRangeBetweenHeaderAndTotal();
    const colSpendERPSek = tablePD.getColumnByName("SPEND - ERP SEK").getRangeBetweenHeaderAndTotal();
    const colUMERP = tablePD.getColumnByName("MEASURE - ERP").getRangeBetweenHeaderAndTotal();
    const colSupERP = tablePD.getColumnByName("SUPPLIER - ERP").getRangeBetweenHeaderAndTotal();
    const colData = tablePD.getColumnByName("DATA FATURAMENTO").getRangeBetweenHeaderAndTotal();

    const colRG = tablePD.getColumnByName("REGIONAL / GLOBAL").getRangeBetweenHeaderAndTotal();
    const colLocalOver = tablePD.getColumnByName("LOCAL /OVERSEA").getRangeBetweenHeaderAndTotal();

    // Leitura dos valores da PD (uma vez)
    const arrArticlevals = colArticle.getValues();
    const arrArticleSCvals = colArticleSC.getValues();
    const arrSupIDvals = colSupID.getValues();
    const arrPriceCtrlvals = colPriceCtrl.getValues();
    const arrCurCtrlvals = colCurCtrl.getValues();
    const arrVolvals = colVol.getValues();
    const arrDatavals = colData.getValues();

    const n = arrArticlevals.length;

    // Índices do mês alvo
    const idxMes: number[] = [];
    for (let i = 0; i < n; i++) {
        const serial = getSerialFromCellValue(arrDatavals[i][0] as unknown);
        if (serial !== null && serial === mesSerial) idxMes.push(i);
    }
    if (idxMes.length === 0) {
        if (hadTotalsPD) tablePD.setShowTotals(true);
        return "Nenhum registro encontrado para o mês informado.";
    }

    countZero = 0;
    countTotal = idxMes.length;

    // Buffers de saída (clonamos para preservar outros meses)
    const outSekCtrl: NumMatrix = cloneValues(colSekCtrl) as NumMatrix;
    const outSpendAct: NumMatrix = cloneValues(colSpendAct) as NumMatrix;
    const outSpendActSek: NumMatrix = cloneValues(colSpendActSek) as NumMatrix;
    const outMeasureActual: StrMatrix = cloneValues(colMeasureActual) as StrMatrix;
    const outCM: StrMatrix = cloneValues(colBuyer) as StrMatrix;

    const outDiffAbs: NumMatrix = cloneValues(colDiffAbs) as NumMatrix;
    const outDiffRel: NumMatrix = cloneValues(colDiffRel) as NumMatrix;

    const outPriceERP: NumMatrix = cloneValues(colPriceERP) as NumMatrix;
    const outCurERP: StrMatrix = cloneValues(colCurERP) as StrMatrix;
    const outSupERP: NumMatrix = cloneValues(colSupERP) as NumMatrix;
    const outUMERP: StrMatrix = cloneValues(colUMERP) as StrMatrix;
    const outPricePCERP: NumMatrix = cloneValues(colPricePCERP) as NumMatrix;

    const outSekERP: NumMatrix = cloneValues(colSekERP) as NumMatrix;
    const outSpendERP: NumMatrix = cloneValues(colSpendERP) as NumMatrix;
    const outSpendERPSek: NumMatrix = cloneValues(colSpendERPSek) as NumMatrix;

    const outRG: StrMatrix = cloneValues(colRG) as StrMatrix;
    const outLocalOver: StrMatrix = cloneValues(colLocalOver) as StrMatrix;

    // Para pintar fallback por Article
    const fallbackYellowIdx: number[] = [];

    const uniqueArticles = new Set<string>();

    const cmMap = new Map<number, string>();

    for (let i = 0; i < refIDvals.length; i++) {
        const supRef = toNumberSafe(refIDvals[i][0]);   // supplier ID
        const cmRef = toStringSafe(refCMvals[i][0]);   // CM

        if (supRef !== null && cmRef) {
            cmMap.set(supRef, cmRef);
        }
    }

    const refMap = new Map<string, string>();

    for (let i = 0; i < refArticlevals.length; i++) {
        const ArticleRef = toUpperCode(refArticlevals[i][0]);   // Article da tabela Refs
        const rgRef = toStringSafe(refRGvals[i][0]);  // "R" ou "G"

        if (ArticleRef && rgRef) {
            refMap.set(ArticleRef, rgRef);
        }
    }

    // ===== Processamento apenas para as linhas do mês =====
    for (const i of idxMes) {
        const ArticleSC = toUpperCode(arrArticleSCvals[i][0]);
        const Article = toUpperCode(arrArticlevals[i][0] as unknown);
        const supId = toNumberSafe(arrSupIDvals[i][0] as unknown);
        const priceCtrl = toNumberSafe(arrPriceCtrlvals[i][0] as unknown);
        const curCtrl = toUpperCode(arrCurCtrlvals[i][0] as unknown);
        const vol = toNumberSafe(arrVolvals[i][0] as unknown);

        // === LOCAL (L) ou OVERSEA (O) ===
        let localOver: string = "";

        if (curCtrl === "BRL") {
            localOver = "L";
        } else {
            localOver = "O";
        }

        outLocalOver[i][0] = localOver;

        // --- ACTUAL SEK ---
        const rateCtrl = ratesMap.get(curCtrl);
        if (rateCtrl !== undefined && priceCtrl !== null) {
            outSekCtrl[i][0] = priceCtrl * rateCtrl;
        }

        // --- Spend / Spend SEK (ACTUAL) ---
        if (vol !== null && priceCtrl !== null) {
            const spend = priceCtrl * vol;
            if (spend !== 0) outSpendAct[i][0] = spend;
            const sek = outSekCtrl[i][0];
            if (sek !== null) {
                const spendSek = sek * vol;
                if (spendSek !== 0) outSpendActSek[i][0] = spendSek;
            }
        }

        // --- ERP (match Article+SUP ou fallback por Article) ---
        let ERP: ERPItem | undefined;
        if (supId !== null && Number.isFinite(supId)) {
            ERP = ERPByArticleSup.get(`${Article}|${supId}`);
        }
        if (!ERP) {
            ERP = ERPByArticle.get(Article);
            if (ERP) fallbackYellowIdx.push(i); // fallback por Article
        }
        if (!ERP) continue;

        outPriceERP[i][0] = Number.isFinite(ERP.price) ? ERP.price : null;
        outCurERP[i][0] = ERP.currency || null;
        outSupERP[i][0] = Number.isFinite(ERP.supplierId) ? ERP.supplierId : null;
        outUMERP[i][0] = ERP.measure || null;

        let pricePC = Number.isFinite(ERP.price) ? ERP.price : null;
        if (pricePC !== null) {
            if (ERP.measure === "MI") pricePC = ERP.price / 1000;
            else if (ERP.measure === "CT") pricePC = ERP.price / 100;
        }
        outPricePCERP[i][0] = pricePC;

        // ERP SEK
        const rateERP = ratesMap.get(ERP.currency);
        if (rateERP !== undefined && Number.isFinite(ERP.price)) {
            outSekERP[i][0] = ERP.price * rateERP;
        }

        // --- Detectar UM do ACTUAL comparando com ERP ---

        const priceActual = priceCtrl;

        // Preços ERP ajustados
        const priceERP_raw = Number.isFinite(ERP.price) ? ERP.price : null;

        let priceERP_PC: number | null = null;
        if (priceERP_raw !== null) {
            if (ERP.measure === "MI") priceERP_PC = priceERP_raw / 1000;
            else if (ERP.measure === "CT") priceERP_PC = priceERP_raw / 100;
            else priceERP_PC = priceERP_raw;
        }

        // Comparações
        let umDetected = "?";
        let ERPPriceAligned: number | null = null;

        if (priceERP_raw !== null && priceActual !== null) {
            const tol: number = Math.max(Math.abs(priceActual) * 0.07, 0.01);

            const diffPC = Math.abs(priceActual - (priceERP_PC ?? Infinity));
            const diffCT = Math.abs(priceActual - (priceERP_raw / 100));
            const diffMI = Math.abs(priceActual - (priceERP_raw / 1000));

            if (diffPC < diffCT && diffPC < diffMI && diffPC < tol) {
                umDetected = "PC";
                ERPPriceAligned = priceERP_PC;
            } else if (diffCT < diffPC && diffCT < diffMI && diffCT < tol) {
                umDetected = "CT";
                ERPPriceAligned = priceERP_raw / 100;
            } else if (diffMI < diffPC && diffMI < diffCT && diffMI < tol) {
                umDetected = "MI";
                ERPPriceAligned = priceERP_raw / 1000;
            } else {
                // fallback: usa PC original do ERP
                umDetected = "PC";
                ERPPriceAligned = priceERP_PC;
            }
        }

        outMeasureActual[i][0] = umDetected;

        // ===== SPEND ERP usando UM alinhada =====
        if (ERPPriceAligned !== null && vol !== null) {
            const spendJAligned = ERPPriceAligned * vol;
            outSpendERP[i][0] = spendJAligned;

            const rateERP = ratesMap.get(ERP.currency);
            if (rateERP !== undefined) {
                outSpendERPSek[i][0] = spendJAligned * rateERP;
            }
        }

        const sekActual = outSekCtrl[i][0];

        // SEK ERP usando PREÇO ALINHADO
        const sekERPaligned = (ERPPriceAligned !== null && rateERP !== undefined)
            ? ERPPriceAligned * rateERP
            : null;

        // Guardar para escrita
        outSekERP[i][0] = sekERPaligned;

        if (sekActual !== null && sekERPaligned !== null) {
            outDiffAbs[i][0] = Math.abs(sekActual - sekERPaligned);

            if (sekActual !== 0) {
                outDiffRel[i][0] = Math.abs((sekActual - sekERPaligned) / sekActual);
            }
        }

        if (outDiffRel[i][0] > maiorDiffRel) {
            maiorDiffRel = outDiffRel[i][0];
        }

        if (outDiffAbs[i][0] > maiorDiffAbs) {
            maiorDiffAbs = outDiffAbs[i][0];
        }

        const ArticleNorm = ArticleSC;  // já tratado com toUpperCode antes
        const rgVal = refMap.get(ArticleNorm);

        if (rgVal) {
            outRG[i][0] = rgVal;    // grava “R” ou “G”
        } else {
            outRG[i][0] = null;     // não encontrado
        }

        // === CM (Category Manager) por Supplier ===
        let cmVal: string | undefined = undefined;

        if (supId !== null) {
            cmVal = cmMap.get(supId);
        }

        if (cmVal) {
            outCM[i][0] = cmVal;
        } else {
            outCM[i][0] = null; // se não encontrar
        }

        // Contabiliza diferenças iguais a 0
        const diffPerc = outDiffRel[i][0];

        if (diffPerc !== null && diffPerc >= 0 && diffPerc <= 0.02 && !uniqueArticles.has(Article)) {
            countZero++;
        }

        if (!uniqueArticles.has(Article)) {
            uniqueArticles.add(Article);
        }

    }

    // ===== Escrita em bloco =====
    colSekCtrl.setValues(outSekCtrl as Matrix);
    colSpendAct.setValues(outSpendAct as Matrix);
    colSpendActSek.setValues(outSpendActSek as Matrix);
    colBuyer.setValues(outCM as Matrix);

    colDiffAbs.setValues(outDiffAbs as Matrix);
    colDiffRel.setValues(outDiffRel as Matrix);

    colPriceERP.setValues(outPriceERP as Matrix);
    colCurERP.setValues(outCurERP as Matrix);
    colSupERP.setValues(outSupERP as Matrix);
    colUMERP.setValues(outUMERP as Matrix);
    colPricePCERP.setValues(outPricePCERP as Matrix);
    colMeasureActual.setValues(outMeasureActual as Matrix);
    colSekERP.setValues(outSekERP as Matrix);
    colSpendERP.setValues(outSpendERP as Matrix);
    colSpendERPSek.setValues(outSpendERPSek as Matrix);

    colRG.setValues(outRG as Matrix);
    colLocalOver.setValues(outLocalOver as Matrix);


    // ===== Pintar fallback por Article (no final) =====
    if (fallbackYellowIdx.length > 0) {
        for (const i of fallbackYellowIdx) {
            colPriceERP.getCell(i, 0).getFormat().getFill().setColor("yellow");
        }
    }

    // Reativar Totals se estava on
    if (hadTotalsPD) tablePD.setShowTotals(true);

    if (countTotal > 0) {
        const percentZero = (countZero / uniqueArticles.size) * 100;
        console.log(`A planilha indica que tivemos ${percentZero.toFixed(2)}% de acuracidade no último mês.`);
        console.log(uniqueArticles, countZero);
        return `A planilha indica que tivemos ${percentZero.toFixed(2)}% de acuracidade no último mês.`;
    } else {
        return `Não houve registros para o mês informado, portanto não foi possível calcular a acuracidade.`;
    }
}