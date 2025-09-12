import * as XLSX from 'xlsx';
import {
  SHEET_NAMES,
  HEADER_STYLE,
  DATA_STYLE,
  NUMBER_STYLE,
  TEXT_STYLE,
  DATE_STYLE,
  TITLE_STYLE,
  SUBTITLE_STYLE,
  RENDICION_COLUMNS,
  COMPRAS_COLUMNS,
  INVENTARIO_COLUMNS,
  TIPOS_COMPRA,
  applyCellStyle,
  setColumnWidths,
  INSTITUTIONAL_COLORS
} from './excelConfig';

export interface ExcelData {
  organizacion: string;
  mes: string;
  año: string;
  fechaElaboracion: string;
  saldoInicialMes: number;
  fondoRotatorioFijo: number;
}

export const createExcelFile = (data: ExcelData) => {
  const wb = XLSX.utils.book_new();
  
  // Crear todas las hojas
  createPortadaSheet(wb, data);
  createRendicionMensualSheet(wb, data);
  createFlujoCajaSheet(wb, data);
  createRegistroComprasSheet(wb, data);
  createInventarioSheet(wb, data);
  createSolicitudReposicionSheet(wb, data);
  
  return wb;
};

// Hoja 1: Portada
const createPortadaSheet = (wb: XLSX.WorkBook, data: ExcelData) => {
  const ws = XLSX.utils.aoa_to_sheet([]);
  
  // Logo placeholder y título principal
  XLSX.utils.sheet_add_aoa(ws, [
    [''],
    [''],
    ['SOCIEDAD DE BENEFICENCIA DE PUNO'],
    [''],
    ['Anexo Práctico de Formatos'],
    ['Reglamento Interno de Negocios'],
    [''],
    [`Mes: ${data.mes} ${data.año}`],
    [`Fecha de elaboración: ${data.fechaElaboracion}`],
    [''],
    ['Documento generado automáticamente'],
    ['Sistema de Gestión Financiera']
  ], { origin: 'A1' });
  
  // Aplicar estilos
  applyCellStyle(ws, 'A3', TITLE_STYLE);
  applyCellStyle(ws, 'A5', SUBTITLE_STYLE);
  applyCellStyle(ws, 'A6', SUBTITLE_STYLE);
  applyCellStyle(ws, 'A8', TEXT_STYLE);
  applyCellStyle(ws, 'A9', TEXT_STYLE);
  
  // Merge celdas para el título
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s: { r: 2, c: 0 }, e: { r: 2, c: 5 } });
  ws['!merges'].push({ s: { r: 4, c: 0 }, e: { r: 4, c: 5 } });
  ws['!merges'].push({ s: { r: 5, c: 0 }, e: { r: 5, c: 5 } });
  
  // Establecer ancho de columnas
  setColumnWidths(ws, [20, 15, 15, 15, 15, 15]);
  
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAMES.PORTADA);
};

// Hoja 2: Rendición Mensual
const createRendicionMensualSheet = (wb: XLSX.WorkBook, data: ExcelData) => {
  const ws = XLSX.utils.aoa_to_sheet([]);
  
  // Encabezados
  const headers = RENDICION_COLUMNS.map(col => col.header);
  XLSX.utils.sheet_add_aoa(ws, [
    [`RENDICIÓN MENSUAL - ${data.mes.toUpperCase()} ${data.año}`],
    [''],
    headers
  ], { origin: 'A1' });
  
  // Saldo inicial
  XLSX.utils.sheet_add_aoa(ws, [
    [1, 'SALDO INICIAL DEL MES', data.fechaElaboracion, '', data.saldoInicialMes, '', data.saldoInicialMes, 'Saldo inicial del período']
  ], { origin: 'A4' });
  
  // Fórmulas para filas adicionales (ejemplo para filas 5-20)
  for (let row = 5; row <= 20; row++) {
    const cellG = `G${row}`;
    const cellE = `E${row}`;
    const cellF = `F${row}`;
    const prevG = `G${row - 1}`;
    
    // Fórmula del saldo: Saldo anterior + Ingreso - Egreso
    ws[cellG] = { f: `${prevG}+${cellE}-${cellF}`, t: 'n' };
  }
  
  // Fila de totales
  const totalRow = 22;
  XLSX.utils.sheet_add_aoa(ws, [
    ['', 'TOTALES', '', '', { f: 'SUM(E4:E21)' }, { f: 'SUM(F4:F21)' }, '', '']
  ], { origin: `A${totalRow}` });
  
  // Aplicar estilos a encabezados
  for (let col = 0; col < headers.length; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: 2, c: col });
    applyCellStyle(ws, cellRef, HEADER_STYLE);
  }
  
  // Aplicar estilos a la fila de totales
  for (let col = 0; col < headers.length; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: totalRow - 1, c: col });
    applyCellStyle(ws, cellRef, HEADER_STYLE);
  }
  
  // Establecer ancho de columnas
  setColumnWidths(ws, RENDICION_COLUMNS.map(col => col.width));
  
  // Merge título
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } });
  
  // Proteger celdas con fórmulas
  for (let row = 4; row <= 21; row++) {
    const cellG = XLSX.utils.encode_cell({ r: row, c: 6 });
    if (ws[cellG]) {
      ws[cellG].s = { ...NUMBER_STYLE, protection: { locked: true } };
    }
  }
  
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAMES.RENDICION);
};

// Hoja 3: Flujo de Caja Mensual
const createFlujoCajaSheet = (wb: XLSX.WorkBook, data: ExcelData) => {
  const ws = XLSX.utils.aoa_to_sheet([]);
  
  XLSX.utils.sheet_add_aoa(ws, [
    [`FLUJO DE CAJA MENSUAL - ${data.mes.toUpperCase()} ${data.año}`],
    [''],
    ['CONCEPTO', 'MONTO (S/.)'],
    ['Saldo inicial', data.saldoInicialMes],
    ['Ingresos del mes', 0], // Actualizar con fórmula: =SUM('Rendición Mensual'.E:E)
    ['Egresos del mes', 0], // Actualizar con fórmula: =SUM('Rendición Mensual'.F:F)
    ['Saldo final', { f: 'B4+B5-B6' }],
    [''],
    ['Fondo rotatorio fijo', data.fondoRotatorioFijo],
    ['Utilidad neta a transferir', { f: 'B7-B9' }],
    [''],
    [''],
    ['INSTRUCCIONES:'],
    ['1. En B5: Escriba =SUM(\'Rendición Mensual\'.E:E)'],
    ['2. En B6: Escriba =SUM(\'Rendición Mensual\'.F:F)'],
    ['3. Las demás fórmulas se calculan automáticamente']
  ], { origin: 'A1' });
  
  // Aplicar estilos
  applyCellStyle(ws, 'A1', TITLE_STYLE);
  applyCellStyle(ws, 'A3', HEADER_STYLE);
  applyCellStyle(ws, 'B3', HEADER_STYLE);
  
  // Aplicar bordes y estilos numéricos
  for (let row = 4; row <= 10; row++) {
    applyCellStyle(ws, `A${row}`, TEXT_STYLE);
    applyCellStyle(ws, `B${row}`, NUMBER_STYLE);
  }
  
  // Destacar resultados importantes
  applyCellStyle(ws, 'A7', { ...TEXT_STYLE, font: { bold: true, color: { rgb: INSTITUTIONAL_COLORS.DARK_BLUE } } });
  applyCellStyle(ws, 'B7', { ...NUMBER_STYLE, font: { bold: true, color: { rgb: INSTITUTIONAL_COLORS.DARK_BLUE } } });
  applyCellStyle(ws, 'A10', { ...TEXT_STYLE, font: { bold: true, color: { rgb: INSTITUTIONAL_COLORS.DARK_BLUE } } });
  applyCellStyle(ws, 'B10', { ...NUMBER_STYLE, font: { bold: true, color: { rgb: INSTITUTIONAL_COLORS.DARK_BLUE } } });
  
  setColumnWidths(ws, [30, 15]);
  
  // Merge título
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } });
  
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAMES.FLUJO_CAJA);
};

// Hoja 4: Registro de Compras
const createRegistroComprasSheet = (wb: XLSX.WorkBook, data: ExcelData) => {
  const ws = XLSX.utils.aoa_to_sheet([]);
  
  const headers = COMPRAS_COLUMNS.map(col => col.header);
  XLSX.utils.sheet_add_aoa(ws, [
    [`REGISTRO DE COMPRAS - ${data.mes.toUpperCase()} ${data.año}`],
    [''],
    headers
  ], { origin: 'A1' });
  
  // Fila de totales
  const totalRow = 25;
  XLSX.utils.sheet_add_aoa(ws, [
    ['', 'TOTALES', '', '', '', { f: 'SUM(F4:F24)' }, '', '']
  ], { origin: `A${totalRow}` });
  
  // Aplicar estilos a encabezados
  for (let col = 0; col < headers.length; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: 2, c: col });
    applyCellStyle(ws, cellRef, HEADER_STYLE);
  }
  
  // Aplicar estilos a la fila de totales
  for (let col = 0; col < headers.length; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: totalRow - 1, c: col });
    applyCellStyle(ws, cellRef, HEADER_STYLE);
  }
  
  setColumnWidths(ws, COMPRAS_COLUMNS.map(col => col.width));
  
  // Merge título
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } });
  
  // Agregar validación de datos para tipo de compra (columna G)
  if (!ws['!datavalidation']) ws['!datavalidation'] = [];
  for (let row = 4; row <= 24; row++) {
    const cellRef = XLSX.utils.encode_cell({ r: row - 1, c: 6 });
    ws['!datavalidation'].push({
      sqref: cellRef,
      type: 'list',
      formula1: `"${TIPOS_COMPRA.join(',')}"`
    });
  }
  
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAMES.COMPRAS);
};

// Hoja 5: Inventario
const createInventarioSheet = (wb: XLSX.WorkBook, data: ExcelData) => {
  const ws = XLSX.utils.aoa_to_sheet([]);
  
  const headers = INVENTARIO_COLUMNS.map(col => col.header);
  XLSX.utils.sheet_add_aoa(ws, [
    [`INVENTARIO - ${data.mes.toUpperCase()} ${data.año}`],
    [''],
    headers
  ], { origin: 'A1' });
  
  // Fórmulas para stock final (ejemplo para filas 4-20)
  for (let row = 4; row <= 20; row++) {
    const cellE = `E${row}`;
    const cellB = `B${row}`;
    const cellC = `C${row}`;
    const cellD = `D${row}`;
    
    // Fórmula del stock final: Stock inicial + Ingresos - Ventas
    ws[cellE] = { f: `${cellB}+${cellC}-${cellD}`, t: 'n' };
  }
  
  // Fila de totales
  const totalRow = 22;
  XLSX.utils.sheet_add_aoa(ws, [
    ['TOTALES', { f: 'SUM(B4:B21)' }, { f: 'SUM(C4:C21)' }, { f: 'SUM(D4:D21)' }, { f: 'SUM(E4:E21)' }, '']
  ], { origin: `A${totalRow}` });
  
  // Aplicar estilos a encabezados
  for (let col = 0; col < headers.length; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: 2, c: col });
    applyCellStyle(ws, cellRef, HEADER_STYLE);
  }
  
  // Aplicar estilos a la fila de totales
  for (let col = 0; col < headers.length; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: totalRow - 1, c: col });
    applyCellStyle(ws, cellRef, HEADER_STYLE);
  }
  
  setColumnWidths(ws, INVENTARIO_COLUMNS.map(col => col.width));
  
  // Merge título
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } });
  
  // Proteger celdas con fórmulas
  for (let row = 3; row <= 20; row++) {
    const cellE = XLSX.utils.encode_cell({ r: row, c: 4 });
    if (ws[cellE]) {
      ws[cellE].s = { ...NUMBER_STYLE, protection: { locked: true } };
    }
  }
  
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAMES.INVENTARIO);
};

// Hoja 6: Solicitud de Reposición
const createSolicitudReposicionSheet = (wb: XLSX.WorkBook, data: ExcelData) => {
  const ws = XLSX.utils.aoa_to_sheet([]);
  
  XLSX.utils.sheet_add_aoa(ws, [
    [`SOLICITUD DE REPOSICIÓN - ${data.mes.toUpperCase()} ${data.año}`],
    [''],
    ['DETALLE', 'MONTO (S/.)'],
    ['Compras', 0], // Actualizar manualmente según registro
    ['Servicios', 0], // Valor manual
    ['Otros', 0], // Valor manual
    [''],
    ['Monto total gastado (a reponer)', { f: 'SUM(B4:B6)' }],
    [''],
    ['Saldo actual del fondo', 0], // Valor manual
    [''],
    ['Observaciones:'],
    ['Este monto debe coincidir con el total de egresos de la Rendición Mensual'],
    ['Para uso exclusivo del fondo rotatorio según Reglamento Interno'],
    [''],
    ['INSTRUCCIONES:'],
    ['1. Complete manualmente los montos de Compras, Servicios y Otros'],
    ['2. El total se calcula automáticamente'],
    ['3. Verifique que coincida con egresos totales de Rendición Mensual']
  ], { origin: 'A1' });
  
  // Aplicar estilos
  applyCellStyle(ws, 'A1', TITLE_STYLE);
  applyCellStyle(ws, 'A3', HEADER_STYLE);
  applyCellStyle(ws, 'B3', HEADER_STYLE);
  
  // Aplicar estilos a las filas de datos
  for (let row = 4; row <= 10; row++) {
    applyCellStyle(ws, `A${row}`, TEXT_STYLE);
    applyCellStyle(ws, `B${row}`, NUMBER_STYLE);
  }
  
  // Destacar total
  applyCellStyle(ws, 'A8', { ...TEXT_STYLE, font: { bold: true, color: { rgb: INSTITUTIONAL_COLORS.DARK_BLUE } } });
  applyCellStyle(ws, 'B8', { ...NUMBER_STYLE, font: { bold: true, color: { rgb: INSTITUTIONAL_COLORS.DARK_BLUE } } });
  
  setColumnWidths(ws, [35, 15]);
  
  // Merge título
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } });
  
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAMES.REPOSICION);
};

export const downloadExcelFile = (wb: XLSX.WorkBook, filename: string) => {
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};