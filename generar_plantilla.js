const ExcelJS = require("exceljs");
const path = require("path");

const ARS = '"$" #,##0.00';
const PCT = "0.0%";
const DATE = "dd/mm/yyyy";

const CATEGORIAS_GASTO = [
  "Comida",
  "Transporte",
  "Alquiler",
  "Servicios",
  "Ocio",
  "Salud",
  "Otros",
];

const CATEGORIAS_INGRESO = ["Sueldo", "Extra", "Otros"];

const MESES = [
  { label: "Enero", num: 1 },
  { label: "Febrero", num: 2 },
  { label: "Marzo", num: 3 },
  { label: "Abril", num: 4 },
  { label: "Mayo", num: 5 },
  { label: "Junio", num: 6 },
  { label: "Julio", num: 7 },
  { label: "Agosto", num: 8 },
  { label: "Septiembre", num: 9 },
  { label: "Octubre", num: 10 },
  { label: "Noviembre", num: 11 },
  { label: "Diciembre", num: 12 },
];

function styleHeader(cell) {
  cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F4E79" },
  };
  cell.alignment = { vertical: "middle", horizontal: "center" };
}

function styleTitle(cell) {
  cell.font = { bold: true, size: 16, color: { argb: "FF1F4E79" } };
}

function styleLabel(cell) {
  cell.font = { bold: true, size: 11 };
  cell.alignment = { vertical: "middle" };
}

function styleSection(cell) {
  cell.font = { bold: true, size: 12, color: { argb: "FF1F4E79" } };
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD6E3F0" },
  };
}

function applyBorder(cell) {
  cell.border = {
    top: { style: "thin", color: { argb: "FFB0B0B0" } },
    left: { style: "thin", color: { argb: "FFB0B0B0" } },
    bottom: { style: "thin", color: { argb: "FFB0B0B0" } },
    right: { style: "thin", color: { argb: "FFB0B0B0" } },
  };
}

async function crearPlantilla() {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Plantilla Gastos Viaje";
  wb.created = new Date();

  const wsConfig = wb.addWorksheet("Configuración", {
    views: [{ showGridLines: false }],
  });
  const wsMov = wb.addWorksheet("Movimientos", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  const wsResumen = wb.addWorksheet("Resumen", {
    views: [{ showGridLines: false }],
  });
  const wsMensual = wb.addWorksheet("Resumen mensual", {
    views: [{ showGridLines: false }],
  });

  // ── Configuración ──────────────────────────────────────────
  wsConfig.columns = [
    { width: 28 },
    { width: 18 },
    { width: 4 },
    { width: 22 },
    { width: 4 },
    { width: 22 },
  ];

  wsConfig.mergeCells("A1:B1");
  styleTitle(wsConfig.getCell("A1"));
  wsConfig.getCell("A1").value = "Configuración del ahorro";

  wsConfig.getCell("A3").value = "Meta del viaje";
  styleSection(wsConfig.getCell("A3"));
  wsConfig.getCell("B3").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD6E3F0" },
  };

  const configRows = [
    ["A4", "Monto objetivo (ARS)", "B4", 500000, ARS],
    ["A5", "Fecha objetivo", "B5", new Date(new Date().getFullYear(), 11, 31), DATE],
    ["A6", "Fecha de inicio", "B6", new Date(), DATE],
    ["A7", "Saldo inicial (ARS)", "B7", 0, ARS],
  ];

  for (const [labelCell, label, valueCell, value, numFmt] of configRows) {
    styleLabel(wsConfig.getCell(labelCell));
    wsConfig.getCell(labelCell).value = label;
    wsConfig.getCell(valueCell).value = value;
    wsConfig.getCell(valueCell).numFmt = numFmt;
    applyBorder(wsConfig.getCell(valueCell));
    wsConfig.getCell(valueCell).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF2CC" },
    };
  }

  wsConfig.getCell("A9").value =
    "Completá las celdas en amarillo. El resto del Excel se calcula solo.";
  wsConfig.getCell("A9").font = { italic: true, color: { argb: "FF666666" } };
  wsConfig.mergeCells("A9:B9");

  // Categorías gasto
  wsConfig.getCell("D3").value = "Categorías de gasto";
  styleSection(wsConfig.getCell("D3"));
  CATEGORIAS_GASTO.forEach((cat, i) => {
    const cell = wsConfig.getCell(4 + i, 4);
    cell.value = cat;
    applyBorder(cell);
  });

  // Categorías ingreso
  wsConfig.getCell("F3").value = "Categorías de ingreso";
  styleSection(wsConfig.getCell("F3"));
  CATEGORIAS_INGRESO.forEach((cat, i) => {
    const cell = wsConfig.getCell(4 + i, 6);
    cell.value = cat;
    applyBorder(cell);
  });

  wsConfig.getCell("A13").value = "Cómo usarla";
  styleSection(wsConfig.getCell("A13"));
  wsConfig.mergeCells("A13:B13");

  const tips = [
    "1. Completá monto objetivo, fecha objetivo, fecha de inicio y saldo inicial.",
    "2. Andá a la hoja Movimientos y cargá cada ingreso o gasto (una fila por operación).",
    "3. Usá los desplegables de Tipo y Categoría para no equivocarte.",
    "4. Mirá Resumen para ver cuánto te falta y cuánto ahorrar por mes.",
    "5. Mirá Resumen mensual para ver la evolución mes a mes.",
  ];
  tips.forEach((tip, i) => {
    wsConfig.getCell(14 + i, 1).value = tip;
    wsConfig.mergeCells(14 + i, 1, 14 + i, 2);
  });

  // ── Movimientos ────────────────────────────────────────────
  const headers = ["Fecha", "Tipo", "Categoría", "Descripción", "Monto"];
  const colWidths = [14, 12, 16, 40, 16];

  headers.forEach((h, i) => {
    const cell = wsMov.getCell(1, i + 1);
    cell.value = h;
    styleHeader(cell);
    applyBorder(cell);
    wsMov.getColumn(i + 1).width = colWidths[i];
  });
  wsMov.getRow(1).height = 22;

  // Ejemplo de filas de muestra (3)
  const ejemplos = [
    [new Date(), "Ingreso", "Sueldo", "Sueldo del mes (ejemplo)", 350000],
    [new Date(), "Gasto", "Comida", "Supermercado (ejemplo)", 45000],
    [new Date(), "Gasto", "Transporte", "Subte / colectivo (ejemplo)", 8000],
  ];

  ejemplos.forEach((row, i) => {
    const r = i + 2;
    wsMov.getCell(r, 1).value = row[0];
    wsMov.getCell(r, 1).numFmt = DATE;
    wsMov.getCell(r, 2).value = row[1];
    wsMov.getCell(r, 3).value = row[2];
    wsMov.getCell(r, 4).value = row[3];
    wsMov.getCell(r, 5).value = row[4];
    wsMov.getCell(r, 5).numFmt = ARS;
    for (let c = 1; c <= 5; c++) applyBorder(wsMov.getCell(r, c));
  });

  // Formato para filas vacías (hasta 500)
  for (let r = 5; r <= 500; r++) {
    wsMov.getCell(r, 1).numFmt = DATE;
    wsMov.getCell(r, 5).numFmt = ARS;
  }

  // Validaciones: Tipo
  wsMov.dataValidations.add("B2:B500", {
    type: "list",
    allowBlank: true,
    formulae: ['"Ingreso,Gasto"'],
    showErrorMessage: true,
    errorTitle: "Tipo inválido",
    error: "Elegí Ingreso o Gasto",
  });

  // Validaciones: Categoría (todas las categorías combinadas)
  const todasCats = [...CATEGORIAS_GASTO, ...CATEGORIAS_INGRESO]
    .filter((v, i, a) => a.indexOf(v) === i)
    .join(",");
  wsMov.dataValidations.add("C2:C500", {
    type: "list",
    allowBlank: true,
    formulae: [`"${todasCats}"`],
    showErrorMessage: true,
    errorTitle: "Categoría inválida",
    error: "Elegí una categoría de la lista",
  });

  // ── Resumen ────────────────────────────────────────────────
  wsResumen.columns = [
    { width: 32 },
    { width: 18 },
    { width: 4 },
    { width: 22 },
    { width: 16 },
  ];

  wsResumen.mergeCells("A1:B1");
  styleTitle(wsResumen.getCell("A1"));
  wsResumen.getCell("A1").value = "Resumen del ahorro";

  // KPIs principales
  const kpiLabels = [
    ["A3", "Total ingresos", "B3", '=SUMIF(Movimientos!B:B,"Ingreso",Movimientos!E:E)'],
    ["A4", "Total gastos", "B4", '=SUMIF(Movimientos!B:B,"Gasto",Movimientos!E:E)'],
    [
      "A5",
      "Ahorro actual",
      "B5",
      "='Configuración'!B7+B3-B4",
    ],
    ["A6", "Meta del viaje", "B6", "='Configuración'!B4"],
    ["A7", "Falta para la meta", "B7", "=MAX(0,B6-B5)"],
    ["A8", "% de avance", "B8", '=IF(B6=0,0,MIN(1,B5/B6))'],
  ];

  for (const [lCell, label, vCell, formula] of kpiLabels) {
    styleLabel(wsResumen.getCell(lCell));
    wsResumen.getCell(lCell).value = label;
    wsResumen.getCell(vCell).value = { formula };
    applyBorder(wsResumen.getCell(vCell));
  }

  wsResumen.getCell("B3").numFmt = ARS;
  wsResumen.getCell("B4").numFmt = ARS;
  wsResumen.getCell("B5").numFmt = ARS;
  wsResumen.getCell("B6").numFmt = ARS;
  wsResumen.getCell("B7").numFmt = ARS;
  wsResumen.getCell("B8").numFmt = PCT;

  // Colores semánticos
  wsResumen.getCell("B3").font = { color: { argb: "FF006600" }, bold: true };
  wsResumen.getCell("B4").font = { color: { argb: "FF990000" }, bold: true };
  wsResumen.getCell("B5").font = { color: { argb: "FF1F4E79" }, bold: true, size: 12 };
  wsResumen.getCell("B7").font = { color: { argb: "FFC65911" }, bold: true };

  // Formato condicional % avance (barra de datos)
  wsResumen.addConditionalFormatting({
    ref: "B8",
    rules: [
      {
        type: "dataBar",
        cfvo: [
          { type: "num", value: 0 },
          { type: "num", value: 1 },
        ],
        color: { argb: "FF5B9BD5" },
        showValue: true,
        gradient: true,
      },
    ],
  });

  // Plan de ahorro mensual
  wsResumen.getCell("A10").value = "Plan para llegar a tiempo";
  styleSection(wsResumen.getCell("A10"));
  wsResumen.getCell("B10").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD6E3F0" },
  };

  styleLabel(wsResumen.getCell("A11"));
  wsResumen.getCell("A11").value = "Fecha objetivo";
  wsResumen.getCell("B11").value = { formula: "='Configuración'!B5" };
  wsResumen.getCell("B11").numFmt = DATE;
  applyBorder(wsResumen.getCell("B11"));

  styleLabel(wsResumen.getCell("A12"));
  wsResumen.getCell("A12").value = "Meses restantes (aprox.)";
  // MAX(1, (año_obj-año_hoy)*12 + (mes_obj-mes_hoy))
  wsResumen.getCell("B12").value = {
    formula:
      '=MAX(1,(YEAR(\'Configuración\'!B5)-YEAR(TODAY()))*12+(MONTH(\'Configuración\'!B5)-MONTH(TODAY())))',
  };
  applyBorder(wsResumen.getCell("B12"));

  styleLabel(wsResumen.getCell("A13"));
  wsResumen.getCell("A13").value = "Ahorrar por mes";
  wsResumen.getCell("B13").value = { formula: "=IF(B12=0,B7,B7/B12)" };
  wsResumen.getCell("B13").numFmt = ARS;
  wsResumen.getCell("B13").font = { bold: true, size: 13, color: { argb: "FF1F4E79" } };
  applyBorder(wsResumen.getCell("B13"));
  wsResumen.getCell("B13").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE2EFDA" },
  };

  wsResumen.getCell("A14").value =
    "Si cada mes ahorrás al menos ese monto, llegás a la meta antes de la fecha objetivo.";
  wsResumen.getCell("A14").font = { italic: true, color: { argb: "FF666666" }, size: 10 };
  wsResumen.mergeCells("A14:B14");

  // Gastos por categoría
  wsResumen.getCell("D3").value = "Gastos por categoría";
  styleSection(wsResumen.getCell("D3"));
  wsResumen.getCell("E3").value = "Monto";
  styleSection(wsResumen.getCell("E3"));

  CATEGORIAS_GASTO.forEach((cat, i) => {
    const row = 4 + i;
    wsResumen.getCell(row, 4).value = cat;
    applyBorder(wsResumen.getCell(row, 4));
    wsResumen.getCell(row, 5).value = {
      formula: `SUMIFS(Movimientos!E:E,Movimientos!B:B,"Gasto",Movimientos!C:C,D${row})`,
    };
    wsResumen.getCell(row, 5).numFmt = ARS;
    applyBorder(wsResumen.getCell(row, 5));
  });

  const totalCatRow = 4 + CATEGORIAS_GASTO.length;
  wsResumen.getCell(totalCatRow, 4).value = "Total gastos";
  wsResumen.getCell(totalCatRow, 4).font = { bold: true };
  applyBorder(wsResumen.getCell(totalCatRow, 4));
  wsResumen.getCell(totalCatRow, 5).value = {
    formula: `SUM(E4:E${totalCatRow - 1})`,
  };
  wsResumen.getCell(totalCatRow, 5).numFmt = ARS;
  wsResumen.getCell(totalCatRow, 5).font = { bold: true };
  applyBorder(wsResumen.getCell(totalCatRow, 5));

  // ── Resumen mensual ────────────────────────────────────────
  wsMensual.columns = [
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 14 },
  ];

  wsMensual.mergeCells("A1:D1");
  styleTitle(wsMensual.getCell("A1"));
  wsMensual.getCell("A1").value = "Resumen mensual";

  styleLabel(wsMensual.getCell("A3"));
  wsMensual.getCell("A3").value = "Año a analizar";
  wsMensual.getCell("B3").value = new Date().getFullYear();
  applyBorder(wsMensual.getCell("B3"));
  wsMensual.getCell("B3").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFF2CC" },
  };

  const mensHeaders = ["Mes", "Ingresos", "Gastos", "Ahorro del mes"];
  mensHeaders.forEach((h, i) => {
    const cell = wsMensual.getCell(5, i + 1);
    cell.value = h;
    styleHeader(cell);
    applyBorder(cell);
  });

  MESES.forEach((mes, i) => {
    const row = 6 + i;
    wsMensual.getCell(row, 1).value = mes.label;
    applyBorder(wsMensual.getCell(row, 1));

    // Ingresos del mes: SUMIFS con YEAR y MONTH vía SUMPRODUCT (más compatible)
    // Usamos SUMIFS con fechas de inicio/fin del mes
    // Inicio = DATE($B$3, mes, 1)
    // Fin = EOMONTH(DATE($B$3, mes, 1), 0)
    wsMensual.getCell(row, 2).value = {
      formula: `SUMIFS(Movimientos!E:E,Movimientos!B:B,"Ingreso",Movimientos!A:A,">="&DATE($B$3,${mes.num},1),Movimientos!A:A,"<="&EOMONTH(DATE($B$3,${mes.num},1),0))`,
    };
    wsMensual.getCell(row, 2).numFmt = ARS;
    applyBorder(wsMensual.getCell(row, 2));

    wsMensual.getCell(row, 3).value = {
      formula: `SUMIFS(Movimientos!E:E,Movimientos!B:B,"Gasto",Movimientos!A:A,">="&DATE($B$3,${mes.num},1),Movimientos!A:A,"<="&EOMONTH(DATE($B$3,${mes.num},1),0))`,
    };
    wsMensual.getCell(row, 3).numFmt = ARS;
    applyBorder(wsMensual.getCell(row, 3));

    wsMensual.getCell(row, 4).value = { formula: `B${row}-C${row}` };
    wsMensual.getCell(row, 4).numFmt = ARS;
    applyBorder(wsMensual.getCell(row, 4));
  });

  // Totales anuales
  const totalRow = 18;
  wsMensual.getCell(totalRow, 1).value = "Total año";
  wsMensual.getCell(totalRow, 1).font = { bold: true };
  applyBorder(wsMensual.getCell(totalRow, 1));
  wsMensual.getCell(totalRow, 2).value = { formula: "SUM(B6:B17)" };
  wsMensual.getCell(totalRow, 2).numFmt = ARS;
  wsMensual.getCell(totalRow, 2).font = { bold: true };
  applyBorder(wsMensual.getCell(totalRow, 2));
  wsMensual.getCell(totalRow, 3).value = { formula: "SUM(C6:C17)" };
  wsMensual.getCell(totalRow, 3).numFmt = ARS;
  wsMensual.getCell(totalRow, 3).font = { bold: true };
  applyBorder(wsMensual.getCell(totalRow, 3));
  wsMensual.getCell(totalRow, 4).value = { formula: "SUM(D6:D17)" };
  wsMensual.getCell(totalRow, 4).numFmt = ARS;
  wsMensual.getCell(totalRow, 4).font = { bold: true };
  applyBorder(wsMensual.getCell(totalRow, 4));

  wsMensual.getCell("A20").value =
    "Cambiá el año en la celda amarilla para ver otro período. Los movimientos deben tener fecha del año elegido.";
  wsMensual.getCell("A20").font = { italic: true, color: { argb: "FF666666" }, size: 10 };
  wsMensual.mergeCells("A20:D20");

  // Orden de hojas
  wb.views = [{ activeTab: 0 }];

  const outPath = path.join(__dirname, "Plantilla_Gastos_Viaje.xlsx");
  await wb.xlsx.writeFile(outPath);
  console.log(`Plantilla creada: ${outPath}`);
}

crearPlantilla().catch((err) => {
  console.error(err);
  process.exit(1);
});
