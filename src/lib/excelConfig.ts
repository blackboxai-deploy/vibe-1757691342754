import * as XLSX from 'xlsx';

// Configuración de colores institucionales
export const INSTITUTIONAL_COLORS = {
  DARK_BLUE: '1f4e79',
  GRAY: '595959',
  WHITE: 'ffffff',
  LIGHT_GRAY: 'f2f2f2'
} as const;

// Configuración de estilos para encabezados
export const HEADER_STYLE = {
  font: { bold: true, color: { rgb: INSTITUTIONAL_COLORS.WHITE } },
  fill: { fgColor: { rgb: INSTITUTIONAL_COLORS.DARK_BLUE } },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } }
  }
};

// Configuración de estilos para celdas de datos
export const DATA_STYLE = {
  border: {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } }
  }
};

// Configuración de estilos para números
export const NUMBER_STYLE = {
  ...DATA_STYLE,
  alignment: { horizontal: 'right' },
  numFmt: '#,##0.00'
};

// Configuración de estilos para texto
export const TEXT_STYLE = {
  ...DATA_STYLE,
  alignment: { horizontal: 'left' }
};

// Configuración de estilos para fechas
export const DATE_STYLE = {
  ...DATA_STYLE,
  alignment: { horizontal: 'center' },
  numFmt: 'dd/mm/yyyy'
};

// Configuración de estilos para títulos
export const TITLE_STYLE = {
  font: { bold: true, size: 16, color: { rgb: INSTITUTIONAL_COLORS.DARK_BLUE } },
  alignment: { horizontal: 'center', vertical: 'center' }
};

// Configuración de estilos para subtítulos
export const SUBTITLE_STYLE = {
  font: { bold: true, size: 14, color: { rgb: INSTITUTIONAL_COLORS.GRAY } },
  alignment: { horizontal: 'center', vertical: 'center' }
};

// Definición de las hojas del Excel
export const SHEET_NAMES = {
  PORTADA: 'Portada',
  RENDICION: 'Rendición Mensual',
  FLUJO_CAJA: 'Flujo de Caja Mensual',
  COMPRAS: 'Registro de Compras',
  INVENTARIO: 'Inventario',
  REPOSICION: 'Solicitud de Reposición'
} as const;

// Estructura de columnas para Rendición Mensual
export const RENDICION_COLUMNS = [
  { key: 'numero', header: 'N°', width: 8 },
  { key: 'concepto', header: 'Concepto', width: 30 },
  { key: 'fecha', header: 'Fecha', width: 12 },
  { key: 'documento', header: 'Documento (N° Boleta/Factura)', width: 20 },
  { key: 'ingreso', header: 'Ingreso (S/.)', width: 15 },
  { key: 'egreso', header: 'Egreso (S/.)', width: 15 },
  { key: 'saldo', header: 'Saldo (S/.)', width: 15 },
  { key: 'observaciones', header: 'Observaciones', width: 25 }
];

// Estructura de columnas para Registro de Compras
export const COMPRAS_COLUMNS = [
  { key: 'numero', header: 'N°', width: 8 },
  { key: 'fecha', header: 'Fecha', width: 12 },
  { key: 'proveedor', header: 'Proveedor', width: 25 },
  { key: 'documento', header: 'Documento (N° Factura)', width: 20 },
  { key: 'concepto', header: 'Concepto', width: 30 },
  { key: 'monto', header: 'Monto (S/.)', width: 15 },
  { key: 'tipo', header: 'Tipo de compra', width: 20 },
  { key: 'observaciones', header: 'Observaciones', width: 25 }
];

// Estructura de columnas para Inventario
export const INVENTARIO_COLUMNS = [
  { key: 'producto', header: 'Producto', width: 30 },
  { key: 'stockInicial', header: 'Stock inicial', width: 15 },
  { key: 'ingresos', header: 'Ingresos (compras)', width: 18 },
  { key: 'ventas', header: 'Ventas', width: 15 },
  { key: 'stockFinal', header: 'Stock final', width: 15 },
  { key: 'observaciones', header: 'Observaciones', width: 25 }
];

// Tipos de compra para validación
export const TIPOS_COMPRA = [
  'Menor (Fondo)',
  'Mayor (Abastecimiento)'
];

// Función para aplicar estilos a una celda
export const applyCellStyle = (ws: XLSX.WorkSheet, cellRef: string, style: any) => {
  if (!ws[cellRef]) ws[cellRef] = {};
  ws[cellRef].s = style;
};

// Función para establecer el ancho de columnas
export const setColumnWidths = (ws: XLSX.WorkSheet, widths: number[]) => {
  if (!ws['!cols']) ws['!cols'] = [];
  widths.forEach((width, index) => {
    if (ws['!cols']) {
      ws['!cols'][index] = { wch: width };
    }
  });
};