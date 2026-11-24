import * as XLSX from 'xlsx';
import { CATALOGO_BY_CODIGO, CATALOGO_BY_GRUPO } from './catalogo-mibi';
import { addDaysISO, formatNowDateTime, getTodayISO, recalcularPedido, round2 } from './finanzas';
import type {
  AdvertenciaExcel,
  CellPrimitive,
  EstadoArmado,
  EstadoEnvio,
  LineaPedido,
  MedioPago,
  Pedido,
  ProductoCatalogo,
  ResultadoParseoLista006,
  SheetJSCellObject,
  SheetJSWorkBook,
  SheetJSWorkSheet,
  VarianteCatalogo
} from './types';

export const limpiarTextoInvisible = (s: string): string =>
  s.replace(/[\u200B-\u200D\uFEFF\u00A0]/g, ' ').replace(/\s+/g, ' ').trim();

export const normalizarTexto = (s: string): string =>
  limpiarTextoInvisible(s)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

const COLORES_EXTRA_INVALIDOS = [
  'gris',
  'negro',
  'negra',
  'blanco',
  'blanca',
  'rojo',
  'roja',
  'amarillo',
  'amarilla',
  'naranja',
  'violeta',
  'lila',
  'celeste',
  'bordo',
  'marron'
] as const;

const escapeRegex = (s: string): string =>
  s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');

export const resolverLinea = (
  producto: ProductoCatalogo,
  cantidadTotal: number,
  notaClienteRaw: string,
  idLineaSuffix = ''
): LineaPedido => {
  const notaLimpia = limpiarTextoInvisible(notaClienteRaw);
  const notaNorm = normalizarTexto(notaLimpia);
  const idLinea = `${producto.idGrupo}${idLineaSuffix ? `-${idLineaSuffix}` : ''}`;

  // Caso 1: Producto de variante única -> asignación directa 100% segura
  if (producto.variantes.length === 1) {
    const v = producto.variantes[0]!;
    return {
      idLinea,
      idGrupo: producto.idGrupo,
      nombreProducto: producto.nombre,
      rubro: producto.rubro,
      precioUnitarioNeto: producto.precioNeto,
      cantidadTotal,
      notaCliente: notaLimpia,
      estado: 'OK',
      requiereRevision: false,
      desglose: [{ codigo: v.codigo, etiqueta: v.etiqueta, cantidad: cantidadTotal }]
    };
  }

  // Caso 2: Producto con 2 variantes (talle o color) -> cero suposiciones
  const v0 = producto.variantes[0]!;
  const v1 = producto.variantes[1]!;

  const fallbackRevision = (motivo: string): LineaPedido => ({
    idLinea,
    idGrupo: producto.idGrupo,
    nombreProducto: producto.nombre,
    rubro: producto.rubro,
    precioUnitarioNeto: producto.precioNeto,
    cantidadTotal,
    notaCliente: notaLimpia,
    estado: 'REQUIERE_REVISION',
    requiereRevision: true,
    motivoRevision: motivo,
    desglose: [
      { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: 0 },
      { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: 0 }
    ]
  });

  if (!notaNorm) {
    return fallbackRevision('Sin detalle de talle o color');
  }

  if (/\b(surtid[oa]s?|variad[oa]s?|eleccion|mezclad[oa]s?|cualquiera)\b/.test(notaNorm)) {
    return fallbackRevision('Nota ambigua ("surtido" o similar)');
  }

  const aliasesTodos = new Set([...v0.aliases, ...v1.aliases]);
  for (const colorInv of COLORES_EXTRA_INVALIDOS) {
    const reColor = new RegExp(`\\b${colorInv}\\b`, 'i');
    if (reColor.test(notaNorm) && !aliasesTodos.has(colorInv)) {
      return fallbackRevision(`Color "${colorInv}" no disponible en este modelo`);
    }
  }

  if (/\b(mitad y mitad|50 y 50|50\/50)\b/.test(notaNorm)) {
    if (cantidadTotal % 2 !== 0) {
      return fallbackRevision(`Cantidad impar (${cantidadTotal}) para repartir mitad y mitad`);
    }
    const mitad = cantidadTotal / 2;
    return {
      idLinea,
      idGrupo: producto.idGrupo,
      nombreProducto: producto.nombre,
      rubro: producto.rubro,
      precioUnitarioNeto: producto.precioNeto,
      cantidadTotal,
      notaCliente: notaLimpia,
      estado: 'OK',
      requiereRevision: false,
      desglose: [
        { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: mitad },
        { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: mitad }
      ]
    };
  }

  // Patrón puro "a y b" / "a / b" / "a + b" (solo números simétricos aceptados automáticamente)
  const matchPuro = notaNorm.match(/^(\d+)\s*(?:y|\/|\+|-)\s*(\d+)$/);
  if (matchPuro) {
    const a = parseInt(matchPuro[1] ?? '0', 10);
    const b = parseInt(matchPuro[2] ?? '0', 10);
    if (a + b !== cantidadTotal) {
      return fallbackRevision(`La suma (${a} + ${b}) no coincide con el total (${cantidadTotal})`);
    }
    if (a !== b) {
      return fallbackRevision(`Reparto asimétrico ("${notaLimpia}") sin indicar talle o color`);
    }
    return {
      idLinea,
      idGrupo: producto.idGrupo,
      nombreProducto: producto.nombre,
      rubro: producto.rubro,
      precioUnitarioNeto: producto.precioNeto,
      cantidadTotal,
      notaCliente: notaLimpia,
      estado: 'OK',
      requiereRevision: false,
      desglose: [
        { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: a },
        { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: b }
      ]
    };
  }

  const tieneRangoT1 = /\b(?:de\s*)?1\s*a\s*3\b|\bt1\b|\btalle\s*1\b/.test(notaNorm);
  const tieneRangoT2 = /\b(?:de\s*)?3\s*a\s*6\b|\bt2\b|\btalle\s*2\b/.test(notaNorm);

  const matchRangoConCantidad = notaNorm.match(
    /^(?:de\s*1\s*a\s*3|1\s*a\s*3|t1|talle\s*1)\s*(?::|=|-)?\s*(\d+)$/
  );
  if (matchRangoConCantidad && !tieneRangoT2) {
    const qtyIndicada = parseInt(matchRangoConCantidad[1] ?? '0', 10);
    if (qtyIndicada !== cantidadTotal) {
      return fallbackRevision(`La nota indica ${qtyIndicada} u. para T1 pero el total pedido es ${cantidadTotal}`);
    }
    return {
      idLinea,
      idGrupo: producto.idGrupo,
      nombreProducto: producto.nombre,
      rubro: producto.rubro,
      precioUnitarioNeto: producto.precioNeto,
      cantidadTotal,
      notaCliente: notaLimpia,
      estado: 'OK',
      requiereRevision: false,
      desglose: [
        { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: cantidadTotal },
        { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: 0 }
      ]
    };
  }

  const notaSinRangosEdad = notaNorm
    .replace(/\b(?:de\s*)?1\s*a\s*3(?:\s*anos)?\b/g, ' ')
    .replace(/\b(?:de\s*)?3\s*a\s*6(?:\s*anos)?\b/g, ' ')
    .replace(/\btalle\s*[12]\b/g, ' ')
    .replace(/\bt[12]\b/g, ' ');
  const numerosSueltos = notaSinRangosEdad.match(/\d+/g);

  const aliasesV0SinTalle = v0.aliases.filter((a) => !['1 a 3', '1-3', 't1', 'talle 1'].includes(a));
  const aliasesV1SinTalle = v1.aliases.filter((a) => !['3 a 6', '3-6', 't2', 'talle 2'].includes(a));

  const mencionaV0 =
    tieneRangoT1 || aliasesV0SinTalle.some((al) => new RegExp(`\\b${escapeRegex(al)}\\b`, 'i').test(notaNorm));
  const mencionaV1 =
    tieneRangoT2 || aliasesV1SinTalle.some((al) => new RegExp(`\\b${escapeRegex(al)}\\b`, 'i').test(notaNorm));

  if (!numerosSueltos) {
    if (mencionaV0 && !mencionaV1) {
      return {
        idLinea,
        idGrupo: producto.idGrupo,
        nombreProducto: producto.nombre,
        rubro: producto.rubro,
        precioUnitarioNeto: producto.precioNeto,
        cantidadTotal,
        notaCliente: notaLimpia,
        estado: 'OK',
        requiereRevision: false,
        desglose: [
          { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: cantidadTotal },
          { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: 0 }
        ]
      };
    }
    if (mencionaV1 && !mencionaV0) {
      return {
        idLinea,
        idGrupo: producto.idGrupo,
        nombreProducto: producto.nombre,
        rubro: producto.rubro,
        precioUnitarioNeto: producto.precioNeto,
        cantidadTotal,
        notaCliente: notaLimpia,
        estado: 'OK',
        requiereRevision: false,
        desglose: [
          { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: 0 },
          { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: cantidadTotal }
        ]
      };
    }
  }

  // Parseo explícito por alias + cantidad con límites de palabra estrictos (\b)
  const empiezaConNumero = /^\s*\d+/.test(notaNorm);
  const extraerCantidadVariante = (variante: VarianteCatalogo): number | null => {
    const aliasPattern = variante.aliases.map(escapeRegex).join('|');
    const reColon = new RegExp(`\\b(?:${aliasPattern})\\b\\s*(?::|=|-)\\s*(\\d+)`, 'i');
    const reAntes = new RegExp(`(\\d+)\\s*(?:u\\.?|unid\\.?|de)?\\s*\\b(?:${aliasPattern})\\b`, 'i');
    const reDespues = new RegExp(`\\b(?:${aliasPattern})\\b\\s*(?::|=|-)?\\s*(\\d+)`, 'i');

    const mColon = notaNorm.match(reColon);
    if (mColon && mColon[1]) return parseInt(mColon[1], 10);

    if (empiezaConNumero) {
      const mAntes = notaNorm.match(reAntes);
      if (mAntes && mAntes[1]) return parseInt(mAntes[1], 10);
      const mDespues = notaNorm.match(reDespues);
      if (mDespues && mDespues[1]) return parseInt(mDespues[1], 10);
    } else {
      const mDespues = notaNorm.match(reDespues);
      if (mDespues && mDespues[1]) return parseInt(mDespues[1], 10);
      const mAntes = notaNorm.match(reAntes);
      if (mAntes && mAntes[1]) return parseInt(mAntes[1], 10);
    }
    return null;
  };

  const c0 = extraerCantidadVariante(v0);
  const c1 = extraerCantidadVariante(v1);
  if (c0 !== null || c1 !== null) {
    const q0 = c0 ?? 0;
    const q1 = c1 ?? 0;
    if (q0 + q1 === cantidadTotal && q0 >= 0 && q1 >= 0) {
      return {
        idLinea,
        idGrupo: producto.idGrupo,
        nombreProducto: producto.nombre,
        rubro: producto.rubro,
        precioUnitarioNeto: producto.precioNeto,
        cantidadTotal,
        notaCliente: notaLimpia,
        estado: 'OK',
        requiereRevision: false,
        desglose: [
          { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: q0 },
          { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: q1 }
        ]
      };
    }
    return fallbackRevision(`Las cantidades (${q0} + ${q1}) no suman ${cantidadTotal}`);
  }

  return fallbackRevision(`No se pudo determinar el desglose exacto ("${notaLimpia}")`);
};

export const aplicarResolucionLinea = (
  linea: LineaPedido,
  q0: number,
  q1: number
): LineaPedido | null => {
  const prod = CATALOGO_BY_GRUPO.get(linea.idGrupo);
  if (!prod || prod.variantes.length < 2) return null;
  if (q0 < 0 || q1 < 0 || !Number.isInteger(q0) || !Number.isInteger(q1)) return null;
  if (q0 + q1 !== linea.cantidadTotal) return null;

  const v0 = prod.variantes[0]!;
  const v1 = prod.variantes[1]!;
  return {
    ...linea,
    estado: 'OK',
    requiereRevision: false,
    motivoRevision: undefined,
    desglose: [
      { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: q0 },
      { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: q1 }
    ]
  };
};

export const crearLineasEjemploLista006 = (): LineaPedido[] => {
  const items: Array<{ codigo: string; qty: number; nota: string }> = [
    { codigo: '6502', qty: 3, nota: '' }, // Chaleco Monster: frenado sin color
    { codigo: '6504', qty: 3, nota: 'de 1 a 3' },
    { codigo: '6564', qty: 6, nota: '3 y 3' },
    { codigo: '6566', qty: 4, nota: '2 y 2' },
    { codigo: '6540', qty: 4, nota: '2 y 2' },
    { codigo: '6541', qty: 4, nota: '2 y 2' },
    { codigo: '6539', qty: 4, nota: '' },
    { codigo: '6542', qty: 4, nota: '' },
    { codigo: '6543', qty: 4, nota: '' },
    { codigo: '6545', qty: 4, nota: '' },
    { codigo: '6510', qty: 3, nota: '' },
    { codigo: '6511', qty: 3, nota: '' },
    { codigo: '6512', qty: 3, nota: '' },
    { codigo: '6513', qty: 3, nota: '' },
    { codigo: '6528', qty: 3, nota: '' },
    { codigo: '6530', qty: 3, nota: '' },
    { codigo: '6594', qty: 1, nota: '' },
    { codigo: '6595', qty: 1, nota: '' },
    { codigo: '6570', qty: 3, nota: '' },
    { codigo: '6575', qty: 3, nota: '' },
    { codigo: '6576', qty: 4, nota: '2 y 2' },
    { codigo: '6586', qty: 4, nota: '2 y 2' },
    { codigo: '6588', qty: 2, nota: '' },
    { codigo: '6589', qty: 2, nota: '' },
    { codigo: '6590', qty: 2, nota: '' },
    { codigo: '6591', qty: 2, nota: '' },
    { codigo: '6593', qty: 2, nota: '' },
    { codigo: '6602', qty: 2, nota: '' },
    { codigo: '6603', qty: 2, nota: '' },
    { codigo: '6604', qty: 2, nota: '' },
    { codigo: '6605', qty: 2, nota: '' },
    { codigo: '6606', qty: 2, nota: '' },
    { codigo: '6607', qty: 2, nota: '' }
  ];

  const lineas: LineaPedido[] = [];
  for (const item of items) {
    const prod = CATALOGO_BY_CODIGO.get(item.codigo);
    if (prod) {
      lineas.push(resolverLinea(prod, item.qty, item.nota));
    }
  }
  return lineas;
};

export const expandirCeldasCombinadas = (ws: SheetJSWorkSheet): CellPrimitive[][] => {
  const ref = ws['!ref'];
  if (!ref) return [];
  const range = XLSX.utils.decode_range(ref);
  const rows: CellPrimitive[][] = [];
  for (let r = 0; r <= range.e.r; r++) {
    const row: CellPrimitive[] = [];
    for (let c = 0; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr] as SheetJSCellObject | undefined;
      if (cell && cell.v !== undefined) {
        if (cell.v instanceof Date) {
          row.push(cell.v.toISOString().slice(0, 10));
        } else if (typeof cell.v === 'string') {
          row.push(limpiarTextoInvisible(cell.v));
        } else {
          row.push(cell.v);
        }
      } else {
        row.push(null);
      }
    }
    rows.push(row);
  }

  const merges = ws['!merges'] ?? [];
  for (const m of merges) {
    const topLeft = rows[m.s.r]?.[m.s.c] ?? null;
    if (topLeft === null || topLeft === '') continue;
    for (let r = m.s.r; r <= m.e.r; r++) {
      const targetRow = rows[r];
      if (!targetRow) continue;
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (targetRow[c] === null || targetRow[c] === '') {
          targetRow[c] = topLeft;
        }
      }
    }
  }
  return rows;
};

const parseCsvLine = (line: string, delimiter: string): string[] => {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]!;
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === delimiter && !inQuotes) {
      result.push(limpiarTextoInvisible(current));
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(limpiarTextoInvisible(current));
  return result;
};

export const leerWorkbookSeguro = (buffer: ArrayBuffer, fileName: string): SheetJSWorkBook => {
  const esCsv = /\.csv$/i.test(fileName);
  if (esCsv) {
    const bytes = new Uint8Array(buffer);
    let texto = new TextDecoder('utf-8').decode(bytes);
    if (texto.includes('\uFFFD')) {
      texto = new TextDecoder('iso-8859-1').decode(bytes);
    }
    texto = texto.replace(/^\uFEFF/, '').trim();
    const lines = texto.split(/\r?\n/);
    const firstNonEmpty = lines.find((l) => l.trim().length > 0) ?? '';
    const semiCount = (firstNonEmpty.match(/;/g) ?? []).length;
    const commaCount = (firstNonEmpty.match(/,/g) ?? []).length;
    const delimiter = semiCount >= commaCount ? ';' : ',';
    const lineasCsv = lines.map((line) => parseCsvLine(line, delimiter));
    const ws = XLSX.utils.aoa_to_sheet(lineasCsv) as unknown as SheetJSWorkSheet;
    return { SheetNames: ['CSV'], Sheets: { CSV: ws } };
  }
  return XLSX.read(buffer, { type: 'array' }) as unknown as SheetJSWorkBook;
};

export const parsearCantidadEstricta = (
  raw: CellPrimitive
): { valid: boolean; qty: number; motivo?: string } => {
  if (raw === null || raw === undefined) return { valid: true, qty: 0 };
  if (typeof raw === 'number') {
    if (!Number.isFinite(raw)) return { valid: false, qty: 0, motivo: 'Número inválido' };
    if (raw === 0) return { valid: true, qty: 0 };
    if (raw < 0) return { valid: false, qty: 0, motivo: `Cantidad negativa (${raw})` };
    if (!Number.isInteger(raw)) return { valid: false, qty: 0, motivo: `Cantidad decimal (${raw})` };
    return { valid: true, qty: raw };
  }
  const str = limpiarTextoInvisible(String(raw));
  if (!str) return { valid: true, qty: 0 };
  if (/^-?\d+$/.test(str)) {
    const n = parseInt(str, 10);
    if (n < 0) return { valid: false, qty: 0, motivo: `Cantidad negativa ("${str}")` };
    return { valid: true, qty: n };
  }
  return { valid: false, qty: 0, motivo: `Cantidad inválida ("${str}")` };
};

export const parsearExcelLista006 = (
  buffer: ArrayBuffer,
  fileName: string
): ResultadoParseoLista006 => {
  let wb: SheetJSWorkBook;
  try {
    wb = leerWorkbookSeguro(buffer, fileName);
  } catch {
    return {
      ok: false,
      errorFatal: 'No se pudo leer el archivo. Verificá que sea un Excel (.xlsx, .xls) o CSV válido.',
      lineas: [],
      advertencias: []
    };
  }

  let esLista006 = false;
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = expandirCeldasCombinadas(ws);
    for (const row of rows) {
      for (const cell of row) {
        if (cell === null) continue;
        const txt = normalizarTexto(String(cell)).replace(/["°º]/g, ' ');
        if (/lista\s*(?:n\s*)?0*6\b/.test(txt)) {
          esLista006 = true;
          break;
        }
      }
      if (esLista006) break;
    }
    if (esLista006) break;
  }

  if (!esLista006) {
    return {
      ok: false,
      errorFatal:
        'El archivo subido no corresponde a la plantilla oficial LISTA N° 006 de Mibi® (no se detectó la cabecera "LISTA N° 006").',
      lineas: [],
      advertencias: []
    };
  }

  const advertencias: AdvertenciaExcel[] = [];
  const lineasMap = new Map<string, LineaPedido>();

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = expandirCeldasCombinadas(ws);

    let colCodeA = 0;
    let colCodeB = 1;
    let colPrecio = 3;
    let colQty = 4;
    let colNota = 6;
    let headerDetectado = false;

    for (let rIdx = 0; rIdx < rows.length; rIdx++) {
      const row = rows[rIdx];
      if (!row) continue;
      const rowTexts = row.map((c) => (c === null ? '' : normalizarTexto(String(c))));

      const idxCantPed = rowTexts.findIndex(
        (t) => t.includes('cantidades pedidas') || t.includes('cantidad pedida')
      );
      const idxCantFallback = rowTexts.findIndex((t) => t === 'cantidad' || t === 'cant');
      const idxTalleColor = rowTexts.findIndex(
        (t) => t.includes('talles/colores') || t.includes('talles / colores') || t.includes('talles/ colores')
      );
      const idxDetalleFallback = rowTexts.findIndex((t) => t.includes('detalle') && !t.includes('producto'));
      const idxCodigos = rowTexts.findIndex((t) => t === 'codigos' || t === 'codigo');
      const idxPrecio = rowTexts.findIndex((t) => t === 'precio' || t.includes('precio neto'));

      if (idxCantPed !== -1 || (idxCodigos !== -1 && (idxCantFallback !== -1 || idxPrecio !== -1))) {
        headerDetectado = true;
        if (idxCodigos !== -1) {
          colCodeA = idxCodigos;
          colCodeB = idxCodigos + 1;
        }
        if (idxPrecio !== -1) colPrecio = idxPrecio;
        if (idxCantPed !== -1) colQty = idxCantPed;
        else if (idxCantFallback !== -1) colQty = idxCantFallback;

        if (idxTalleColor !== -1) colNota = idxTalleColor;
        else if (idxDetalleFallback !== -1) colNota = idxDetalleFallback;
        continue;
      }

      const rawA = row[colCodeA] !== null && row[colCodeA] !== undefined ? String(row[colCodeA]) : '';
      const rawB =
        colCodeB < row.length && row[colCodeB] !== null && row[colCodeB] !== undefined
          ? String(row[colCodeB])
          : '';
      const matchA = limpiarTextoInvisible(rawA).match(/\b(\d{4})\b/);
      const matchB = limpiarTextoInvisible(rawB).match(/\b(\d{4})\b/);

      const codigosEnFila: string[] = [];
      if (matchA && matchA[1]) codigosEnFila.push(matchA[1]);
      if (matchB && matchB[1] && matchB[1] !== matchA?.[1]) codigosEnFila.push(matchB[1]);

      if (codigosEnFila.length === 0) continue;

      const qtyCell = headerDetectado ? (row[colQty] ?? null) : (row[4] ?? null);
      const precioCell = headerDetectado ? (row[colPrecio] ?? null) : (row[3] ?? null);
      const notaCell = headerDetectado ? (row[colNota] ?? null) : (row[6] ?? null);

      // Regla oficial: ignorar fila duplicada 6545 sin precio ni cantidad
      if (
        codigosEnFila.includes('6545') &&
        (precioCell === null || String(precioCell).trim() === '') &&
        (qtyCell === null || String(qtyCell).trim() === '' || Number(qtyCell) === 0)
      ) {
        continue;
      }

      const qtyRes = parsearCantidadEstricta(qtyCell);
      if (!qtyRes.valid) {
        advertencias.push({
          hoja: sheetName,
          fila: rIdx + 1,
          codigo: codigosEnFila.join('/'),
          mensaje: qtyRes.motivo ?? 'Cantidad inválida'
        });
        continue;
      }

      if (qtyRes.qty === 0) continue;

      let producto: ProductoCatalogo | undefined;
      for (const cod of codigosEnFila) {
        const encontrado = CATALOGO_BY_CODIGO.get(cod);
        if (encontrado) {
          producto = encontrado;
          break;
        }
      }

      if (!producto) {
        advertencias.push({
          hoja: sheetName,
          fila: rIdx + 1,
          codigo: codigosEnFila.join('/'),
          mensaje: `Código desconocido (${codigosEnFila.join('/')}) con cantidad ${qtyRes.qty}`
        });
        continue;
      }

      const notaStr = notaCell !== null && notaCell !== undefined ? String(notaCell) : '';
      const linea = resolverLinea(producto, qtyRes.qty, notaStr);
      const existente = lineasMap.get(producto.idGrupo);
      if (existente) {
        const suma = existente.cantidadTotal + linea.cantidadTotal;
        const notaCombinada = [existente.notaCliente, linea.notaCliente].filter(Boolean).join(' · ');
        lineasMap.set(producto.idGrupo, resolverLinea(producto, suma, notaCombinada));
      } else {
        lineasMap.set(producto.idGrupo, linea);
      }
    }
  }

  return {
    ok: true,
    lineas: Array.from(lineasMap.values()),
    advertencias
  };
};

export const parsearMontoArgentino = (raw: CellPrimitive): number => {
  if (raw === null || raw === undefined) return 0;
  if (typeof raw === 'number') return Number.isFinite(raw) ? round2(raw) : 0;
  const s = limpiarTextoInvisible(String(raw));
  if (!s || /^#VALUE!$/i.test(s)) return 0;
  if (/^\d{1,2}\/\d{1,2}(?:\/\d{2,4})?$/.test(s)) return 0;

  const sinSimboloRaw = s.replace(/\$/g, '').trim();
  const tokenNum = sinSimboloRaw.match(/^(-?\d[\d.,]*)/)?.[1] ?? '';
  const sinSimbolo = tokenNum || sinSimboloRaw;
  if (/^\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?$/.test(sinSimbolo)) {
    return round2(parseFloat(sinSimbolo.replace(/\./g, '').replace(',', '.')));
  }
  if (/^\d{1,3}(?:,\d{3})+,\d{2}$/.test(sinSimbolo)) {
    const idx = sinSimbolo.lastIndexOf(',');
    const limpio = `${sinSimbolo.slice(0, idx).replace(/,/g, '')}.${sinSimbolo.slice(idx + 1)}`;
    return round2(parseFloat(limpio));
  }
  const simple = sinSimbolo.replace(/[^0-9.,-]/g, '');
  if (!simple) return 0;
  if (simple.includes(',') && !simple.includes('.')) {
    return round2(parseFloat(simple.replace(',', '.')));
  }
  return round2(parseFloat(simple.replace(/,/g, '')) || 0);
};

export const parsearFechaExcel = (raw: CellPrimitive, fallbackISO: string): string => {
  if (raw === null || raw === undefined) return fallbackISO;
  if (typeof raw === 'number' && Number.isFinite(raw) && raw > 25000 && raw < 70000) {
    const utcMs = Date.UTC(1899, 11, 30) + Math.round(raw) * 86400000;
    const d = new Date(utcMs);
    let y = d.getUTCFullYear();
    if (y === 2006) y = 2026;
    const m = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  const str = limpiarTextoInvisible(String(raw));
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str.startsWith('2006-') ? `2026-${str.slice(5)}` : str;
  }
  const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?$/);
  if (dmy) {
    const day = (dmy[1] ?? '').padStart(2, '0');
    const month = (dmy[2] ?? '').padStart(2, '0');
    let year = dmy[3] ? (dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3]) : String(new Date().getFullYear());
    if (year === '2006') year = '2026';
    return `${year}-${month}-${day}`;
  }
  return fallbackISO;
};

export const importarPlanillaGeneral = (
  buffer: ArrayBuffer,
  fileName: string,
  existentes: ReadonlyArray<Pick<Pedido, 'numeroPedido' | 'cliente'>>
): { ok: boolean; nuevos: Pedido[]; duplicadosOmitidos: number; error?: string } => {
  let wb: SheetJSWorkBook;
  try {
    wb = leerWorkbookSeguro(buffer, fileName);
  } catch {
    return { ok: false, nuevos: [], duplicadosOmitidos: 0, error: 'No se pudo leer la planilla general.' };
  }

  const sheetName =
    wb.SheetNames.find((n) => normalizarTexto(n).includes('pedidos mayorista')) ?? wb.SheetNames[0];
  const ws = sheetName ? wb.Sheets[sheetName] : undefined;
  if (!ws) {
    return { ok: false, nuevos: [], duplicadosOmitidos: 0, error: 'La planilla no tiene hojas visibles.' };
  }

  const rows = expandirCeldasCombinadas(ws);
  const headerIdx = rows.findIndex((r) => {
    const norms = r.map((c) => (c === null ? '' : normalizarTexto(String(c))));
    const unicos = new Set(norms.filter(Boolean));
    if (unicos.size < 4) return false;
    const tieneCliente = norms.some(
      (t) => t === 'cliente / negocio' || t === 'cliente' || t.startsWith('cliente /')
    );
    const tienePedido = norms.some((t) => t.includes('pedido') || t.includes('total'));
    const tieneArmadoOPago = norms.some(
      (t) => t === 'armado' || t.includes('estado de pago') || t.includes('monto abonado')
    );
    return tieneCliente && tienePedido && tieneArmadoOPago;
  });

  if (headerIdx === -1) {
    return {
      ok: false,
      nuevos: [],
      duplicadosOmitidos: 0,
      error: 'No se detectó la fila de encabezados de PEDIDOS MAYORISTA MIBI.'
    };
  }

  const headerRow = rows[headerIdx] ?? [];
  const headers = headerRow.map((h) => (h === null ? '' : normalizarTexto(String(h))));
  const findCol = (needle: string): number => headers.findIndex((h) => h.includes(needle));
  const findExactOrCol = (candidates: ReadonlyArray<string>): number => {
    for (const cand of candidates) {
      const idx = headers.findIndex((h) => h === cand || h.includes(cand));
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const cNum = findExactOrCol(['n.º pedido', 'n.o pedido', 'n° pedido', 'nro pedido', 'n pedido']);
  const cCli = findExactOrCol(['cliente / negocio', 'cliente']);
  const cVend = findCol('vendedor');
  const cWsp = findCol('whatsapp');
  const cFec = findCol('fecha pedido');
  const cDesp = findCol('despacho previsto');
  const cDet = findCol('detalle');
  const cTot = findCol('total pedido');
  const cAbo = findCol('monto abonado');
  const cSal = findCol('saldo pendiente');
  const cEstPago = findCol('estado de pago');
  const cArm = findCol('armado');
  const cEnv = findCol('envio');
  const cFecDesp = findCol('fecha despacho');
  const cFecEnt = findCol('fecha entrega');
  const cTrans = findCol('transporte');
  const cGuia = findCol('seguimiento');
  const cCheq = findCol('cheque');
  const cObs = findCol('observaciones');

  const clavesExistentes = new Set(
    existentes.map((p) => `${p.numeroPedido.replace(/^0+/, '')}|${normalizarTexto(p.cliente)}`)
  );

  let maxNum = existentes.reduce((acc, p) => {
    const n = parseInt(p.numeroPedido.replace(/\D/g, ''), 10);
    return Number.isFinite(n) && n > acc ? n : acc;
  }, 0);

  const nuevos: Pedido[] = [];
  let duplicadosOmitidos = 0;
  const hoy = getTodayISO();
  const baseTs = Date.now();

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const numCell = cNum !== -1 ? (r[cNum] ?? null) : null;
    const cliCell = cCli !== -1 ? (r[cCli] ?? null) : null;
    const numRaw = numCell !== null ? limpiarTextoInvisible(String(numCell)) : '';
    const cliRaw = cliCell !== null ? limpiarTextoInvisible(String(cliCell)) : '';
    if (!cliRaw || /^total\b/i.test(cliRaw) || /^total\b/i.test(numRaw)) continue;

    let numeroPedido: string;
    if (numRaw) {
      numeroPedido = numRaw.replace(/^#/, '').padStart(4, '0');
      const parsedNum = parseInt(numeroPedido.replace(/\D/g, ''), 10);
      if (Number.isFinite(parsedNum) && parsedNum > maxNum) maxNum = parsedNum;
    } else {
      maxNum += 1;
      numeroPedido = String(maxNum).padStart(4, '0');
    }

    const clave = `${numeroPedido.replace(/^0+/, '')}|${normalizarTexto(cliRaw)}`;
    if (clavesExistentes.has(clave)) {
      duplicadosOmitidos += 1;
      continue;
    }
    clavesExistentes.add(clave);

    const totalFacturado = parsearMontoArgentino(cTot !== -1 ? (r[cTot] ?? 0) : 0);
    const aboCell = cAbo !== -1 ? (r[cAbo] ?? null) : null;
    const aboVacio = aboCell === null || limpiarTextoInvisible(String(aboCell)) === '';
    let montoAbonado = parsearMontoArgentino(aboCell);
    if (montoAbonado > totalFacturado && totalFacturado > 0) montoAbonado = totalFacturado;

    const salCell = cSal !== -1 ? (r[cSal] ?? null) : null;
    const salStr = salCell !== null ? limpiarTextoInvisible(String(salCell)) : '';
    const saldoSheet = salStr !== '' && !/^#VALUE!$/i.test(salStr) ? parsearMontoArgentino(salCell) : null;

    const estPagoCell = cEstPago !== -1 ? (r[cEstPago] ?? null) : null;
    const estPagoRaw = estPagoCell !== null ? limpiarTextoInvisible(String(estPagoCell)) : '';
    const estPagoStr = normalizarTexto(estPagoRaw);
    const cheqCell = cCheq !== -1 ? (r[cCheq] ?? null) : null;
    const cheqRaw = cheqCell !== null ? limpiarTextoInvisible(String(cheqCell)) : '';
    const cheqStr = normalizarTexto(cheqRaw);

    const esCheque = estPagoStr.includes('cheque') || cheqStr.includes('cheque');
    const chequeAcreditado = esCheque && (estPagoStr.includes('acredit') || estPagoStr === 'pagado');

    if (aboVacio && saldoSheet !== null && totalFacturado > 0 && !esCheque) {
      if (saldoSheet <= 0.01) montoAbonado = totalFacturado;
      else if (saldoSheet < totalFacturado && totalFacturado - saldoSheet > 5) {
        montoAbonado = round2(totalFacturado - saldoSheet);
      }
    }

    let montoCheque = 0;
    if (esCheque && !chequeAcreditado) {
      if (montoAbonado >= totalFacturado && totalFacturado > 0 && (saldoSheet === null || saldoSheet <= 0.01)) {
        montoCheque = totalFacturado;
        montoAbonado = 0;
      } else {
        montoCheque = Math.max(0, round2(totalFacturado - montoAbonado));
      }
    }

    const esPagado =
      (!esCheque && (estPagoStr === 'pagado' || (totalFacturado > 0 && Math.abs(totalFacturado - montoAbonado) <= 0.01))) ||
      chequeAcreditado;

    const medioPago: MedioPago = esCheque
      ? 'Cheque'
      : esPagado || montoAbonado > 0
        ? 'Transferencia / Efectivo'
        : 'Sin pagar';

    const wspCell = cWsp !== -1 ? (r[cWsp] ?? null) : null;
    const wspRaw = wspCell !== null ? limpiarTextoInvisible(String(wspCell)) : '';
    const vendCell = cVend !== -1 ? (r[cVend] ?? null) : null;
    const vendCandidate = vendCell !== null ? limpiarTextoInvisible(String(vendCell)) : '';
    const vendRaw = vendCandidate
      ? vendCandidate
      : !/^\d+$/.test(wspRaw.replace(/[\s\-+]/g, '')) && wspRaw
        ? wspRaw
        : 'Administración';

    const armRaw = normalizarTexto(String(cArm !== -1 ? (r[cArm] ?? '') : ''));
    const estadoArmado: EstadoArmado = armRaw.includes('listo')
      ? 'Listo'
      : armRaw.includes('proceso') || (armRaw.includes('armado') && armRaw.includes('en'))
        ? 'En armado'
        : 'Por armar';

    const envRaw = normalizarTexto(String(cEnv !== -1 ? (r[cEnv] ?? '') : ''));
    const estadoEnvio: EstadoEnvio = envRaw.includes('entreg')
      ? 'Entregado'
      : envRaw.includes('despach')
        ? 'Despachado'
        : 'Pendiente';

    const fechaPedido = parsearFechaExcel(cFec !== -1 ? (r[cFec] ?? null) : null, hoy);
    const despachoPrevisto = parsearFechaExcel(cDesp !== -1 ? (r[cDesp] ?? null) : null, '');
    const fechaDespacho = parsearFechaExcel(cFecDesp !== -1 ? (r[cFecDesp] ?? null) : null, '');
    const fechaEntrega = parsearFechaExcel(cFecEnt !== -1 ? (r[cFecEnt] ?? null) : null, '');

    const ped = recalcularPedido({
      id: `imp_${baseTs}_${i}`,
      numeroPedido,
      cliente: cliRaw,
      vendedor: vendRaw,
      whatsapp: wspRaw,
      fechaPedido,
      despachoPrevisto,
      fechaDespacho: fechaDespacho || undefined,
      fechaEntrega: fechaEntrega || undefined,
      lineas: [],
      descuentoPorc: 0,
      medioPago,
      montoAbonado: esPagado ? totalFacturado : montoAbonado,
      montoCheque,
      chequeAcreditado,
      fechaRecepcionCheque: esCheque ? fechaPedido : '',
      fechaAcreditacionCheque: esCheque ? addDaysISO(fechaPedido, 30) : '',
      detalleCheque: esCheque ? cheqRaw || estPagoRaw : '',
      estadoArmado,
      estadoEnvio,
      transporte: cTrans !== -1 && r[cTrans] ? limpiarTextoInvisible(String(r[cTrans])) : '',
      nroSeguimiento: cGuia !== -1 && r[cGuia] ? limpiarTextoInvisible(String(r[cGuia])) : '',
      observaciones: cObs !== -1 && r[cObs] ? limpiarTextoInvisible(String(r[cObs])) : '',
      version: 1,
      deletedAt: null,
      createdAt: baseTs - i,
      updatedAt: baseTs,
      historial: [
        {
          fechaHora: formatNowDateTime(),
          usuario: 'Importador',
          rol: 'Administración',
          accion: 'Importado desde planilla general'
        }
      ],
      detalleManualImportado: cDet !== -1 && r[cDet] ? limpiarTextoInvisible(String(r[cDet])) : '—',
      totalFacturadoOverride: totalFacturado
    });
    nuevos.push(ped);
  }

  return { ok: true, nuevos, duplicadosOmitidos };
};
