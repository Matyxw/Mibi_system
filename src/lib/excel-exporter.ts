import 'server-only';

import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { MIBI_BRAND } from './catalogo-mibi';
import { limpiarTextoInvisible } from './excel-parser';
import { esDespachoAtrasado, formatShortDate, getTodayISO, round2 } from './finanzas';
import type {
  CellPrimitive,
  Pedido,
  ResultadoPickingConsolidado,
  ResumenFinanciero,
  SheetJSCellObject
} from './types';

const aplicarFormatoMoneda = (ws: XLSX.WorkSheet, cols: ReadonlyArray<number>): void => {
  const ref = ws['!ref'];
  if (!ref) return;
  const range = XLSX.utils.decode_range(ref);
  for (let r = 0; r <= range.e.r; r++) {
    for (const c of cols) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr] as SheetJSCellObject | undefined;
      if (cell && typeof cell.v === 'number') {
        cell.t = 'n';
        cell.z = '"$"#,##0.00';
      }
    }
  }
};

const isoToExcelSerial = (iso: string): number | null => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return null;
  const parts = iso.split('-').map(Number);
  const y = parts[0];
  const m = parts[1];
  const d = parts[2];
  if (!y || !m || !d) return null;
  const utcMs = Date.UTC(y, m - 1, d);
  return Math.round((utcMs - Date.UTC(1899, 11, 30)) / 86400000);
};

const escapeXml = (str: string): string =>
  str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

export interface BufferHojaDepositoResult {
  success: boolean;
  errors: string[];
  buffer?: Uint8Array;
  filename?: string;
}

export const generarBufferHojaDeposito = (pedido: Pedido): BufferHojaDepositoResult => {
  if (pedido.faltantesVariante > 0 || pedido.lineas.some((l) => l.requiereRevision || l.estado !== 'OK')) {
    return {
      success: false,
      errors: [
        `Bloqueado: el pedido #${pedido.numeroPedido} tiene ${pedido.faltantesVariante} artículo(s) sin talle o color definido.`
      ]
    };
  }
  if (pedido.lineas.length === 0) {
    return {
      success: false,
      errors: [`El pedido #${pedido.numeroPedido} no tiene códigos desglosados para depósito.`]
    };
  }

  const filasArticulos: CellPrimitive[][] = [];
  let totalUnidades = 0;
  for (const l of pedido.lineas) {
    for (const d of l.desglose.filter((x) => x.cantidad > 0)) {
      totalUnidades += d.cantidad;
      const subtotalLinea = round2(l.precioUnitarioNeto * d.cantidad);
      filasArticulos.push([
        d.codigo,
        l.nombreProducto,
        d.etiqueta,
        d.cantidad,
        l.precioUnitarioNeto,
        subtotalLinea
      ]);
    }
  }

  const aoa: CellPrimitive[][] = [
    [`MIBI® — HOJA PARA DEPÓSITO · PEDIDO #${pedido.numeroPedido}`],
    [MIBI_BRAND.pieOficial],
    ['Cliente', pedido.cliente, 'Vendedor', pedido.vendedor, 'Fecha pedido', pedido.fechaPedido],
    [
      'Transporte',
      pedido.transporte || 'A definir',
      'Despacho previsto',
      pedido.despachoPrevisto,
      'WhatsApp',
      pedido.whatsapp || '—'
    ],
    [],
    ['Código', 'Artículo', 'Talle / Color', 'Cantidad', 'Precio Neto Unit.', 'Subtotal Neto'],
    ...filasArticulos,
    [],
    ['', '', 'TOTAL UNIDADES', totalUnidades, 'Subtotal Neto', pedido.subtotalNeto],
    ['', '', '', '', `Descuento (${pedido.descuentoPorc}%)`, pedido.montoDescuento],
    ['', '', '', '', 'Base Imponible (Total B)', pedido.baseImponible],
    ['', '', '', '', 'IVA 21%', pedido.iva21],
    ['', '', '', '', 'TOTAL FACTURADO', pedido.totalFacturado],
    ['', '', '', '', 'Cobrado en mano', pedido.montoAbonado],
    ['', '', '', '', 'Saldo pendiente', pedido.saldoPendiente]
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 12 }, { wch: 42 }, { wch: 22 }, { wch: 12 }, { wch: 20 }, { wch: 18 }];
  aplicarFormatoMoneda(ws, [4, 5]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `Pedido_${pedido.numeroPedido}`);
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Uint8Array;
  return {
    success: true,
    errors: [],
    buffer: new Uint8Array(buffer),
    filename: `Mibi_Deposito_Pedido_${pedido.numeroPedido}_${getTodayISO()}.xlsx`
  };
};

export interface BufferPickingResult extends ResultadoPickingConsolidado {
  buffer?: Uint8Array;
  filename?: string;
}

export const generarBufferPickingConsolidado = (
  seleccionados: ReadonlyArray<Pedido>
): BufferPickingResult => {
  const validos: Pedido[] = [];
  const excluidosTalle: string[] = [];
  const excluidosSinLineas: string[] = [];

  for (const p of seleccionados) {
    if (p.faltantesVariante > 0 || p.lineas.some((l) => l.requiereRevision || l.estado !== 'OK')) {
      excluidosTalle.push(`#${p.numeroPedido}`);
    } else if (p.lineas.length === 0) {
      excluidosSinLineas.push(`#${p.numeroPedido}`);
    } else {
      validos.push(p);
    }
  }

  if (validos.length === 0) {
    const motivo =
      excluidosTalle.length > 0
        ? `Frenado: todos los pedidos seleccionados tienen talles o colores sin definir (${excluidosTalle.join(', ')}).`
        : 'Los pedidos seleccionados no tienen desglose de artículos por código.';
    return {
      success: false,
      incluidos: 0,
      excluidosTalle,
      excluidosSinLineas,
      totalUnidades: 0,
      error: motivo
    };
  }

  interface AcumuladoSku {
    codigo: string;
    nombre: string;
    etiqueta: string;
    cantidad: number;
    pedidos: Set<string>;
  }

  const porSku = new Map<string, AcumuladoSku>();
  const filasPorCliente: CellPrimitive[][] = [];
  let totalUnidades = 0;

  for (const p of validos) {
    for (const l of p.lineas) {
      for (const d of l.desglose.filter((x) => x.cantidad > 0)) {
        totalUnidades += d.cantidad;
        const existente = porSku.get(d.codigo);
        if (existente) {
          existente.cantidad += d.cantidad;
          existente.pedidos.add(`#${p.numeroPedido}`);
        } else {
          porSku.set(d.codigo, {
            codigo: d.codigo,
            nombre: l.nombreProducto,
            etiqueta: d.etiqueta,
            cantidad: d.cantidad,
            pedidos: new Set([`#${p.numeroPedido}`])
          });
        }

        filasPorCliente.push([
          `#${p.numeroPedido}`,
          p.cliente,
          p.transporte || 'A definir',
          d.codigo,
          l.nombreProducto,
          d.etiqueta,
          d.cantidad
        ]);
      }
    }
  }

  const filasConsolidadas: CellPrimitive[][] = Array.from(porSku.values())
    .sort((a, b) => a.codigo.localeCompare(b.codigo))
    .map((item) => [
      item.codigo,
      item.nombre,
      item.etiqueta,
      item.cantidad,
      Array.from(item.pedidos).join(', ')
    ]);

  const notaExcluidos =
    excluidosTalle.length > 0
      ? `Excluidos por falta de talle/color: ${excluidosTalle.join(', ')}`
      : 'Todos los códigos validados';

  const hoja1Data: CellPrimitive[][] = [
    ['MIBI® — PICKING CONSOLIDADO PARA DEPÓSITO'],
    [MIBI_BRAND.pieOficial],
    [
      'Fecha emisión',
      getTodayISO(),
      'Pedidos incluidos',
      validos.map((p) => `#${p.numeroPedido}`).join(', '),
      notaExcluidos
    ],
    [],
    ['Código', 'Artículo', 'Talle / Color', 'Total Unidades', 'Pedidos incluidos'],
    ...filasConsolidadas,
    [],
    ['', '', 'TOTAL UNIDADES A RETIRAR', totalUnidades, `${validos.length} pedidos`]
  ];

  const ws1 = XLSX.utils.aoa_to_sheet(hoja1Data);
  ws1['!cols'] = [{ wch: 12 }, { wch: 44 }, { wch: 22 }, { wch: 16 }, { wch: 30 }];

  const hoja2Data: CellPrimitive[][] = [
    ['MIBI® — DESGLOSE POR CLIENTE (SELECCIONADOS)'],
    [MIBI_BRAND.pieOficial],
    [],
    ['N.º pedido', 'Cliente', 'Transporte', 'Código', 'Artículo', 'Talle / Color', 'Cantidad'],
    ...filasPorCliente
  ];

  const ws2 = XLSX.utils.aoa_to_sheet(hoja2Data);
  ws2['!cols'] = [
    { wch: 12 },
    { wch: 26 },
    { wch: 18 },
    { wch: 12 },
    { wch: 42 },
    { wch: 22 },
    { wch: 11 }
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, 'Picking Consolidado');
  XLSX.utils.book_append_sheet(wb, ws2, 'Desglose por Cliente');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Uint8Array;

  return {
    success: true,
    incluidos: validos.length,
    excluidosTalle,
    excluidosSinLineas,
    totalUnidades,
    buffer: new Uint8Array(buffer),
    filename: `Mibi_Picking_Consolidado_${validos.length}ped_${getTodayISO()}.xlsx`
  };
};

export const generarBufferTableroMaestro = async (
  pedidos: ReadonlyArray<Pedido>,
  resumen: ResumenFinanciero,
  sufijoArchivo = ''
): Promise<{ buffer: Uint8Array; filename: string }> => {
  const sufijo = sufijoArchivo ? `_${sufijoArchivo}` : '';
  const nombreArchivo = `PEDIDOS_MAYORISTA_MIBI${sufijo}_${getTodayISO()}.xlsx`;

  const pedidosOrdenados = [...pedidos].sort((a, b) => {
    const na = parseInt(a.numeroPedido.replace(/\D/g, ''), 10);
    const nb = parseInt(b.numeroPedido.replace(/\D/g, ''), 10);
    if (Number.isFinite(na) && Number.isFinite(nb) && na !== nb) return na - nb;
    return a.numeroPedido.localeCompare(b.numeroPedido);
  });

  try {
    const templatePath = path.join(process.cwd(), 'public', 'plantilla_pedidos_mibi.xlsx');
    const tplBuf = await fs.readFile(templatePath);
    const zip = await JSZip.loadAsync(tplBuf);

    const ssFile = zip.file('xl/sharedStrings.xml');
    const s1File = zip.file('xl/worksheets/sheet1.xml');
    const s2File = zip.file('xl/worksheets/sheet2.xml');

    if (ssFile && s1File && s2File) {
      const ssXml = await ssFile.async('string');
      const siList: string[] = [];
      const ssMap = new Map<string, number>();
      for (const m of ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
        siList.push(m[0]);
        const plain = (m[1] ?? '').replace(/<[^>]+>/g, '');
        if (!ssMap.has(plain)) ssMap.set(plain, siList.length - 1);
      }

      const internStr = (val: string): number => {
        const clean = limpiarTextoInvisible(val);
        const existing = ssMap.get(clean);
        if (existing !== undefined) return existing;
        const idx = siList.length;
        const preserve = /^\s|\s$/.test(clean) ? ' xml:space="preserve"' : '';
        siList.push(`<si><t${preserve}>${escapeXml(clean)}</t></si>`);
        ssMap.set(clean, idx);
        return idx;
      };

      const renderStrCell = (ref: string, style: number, text: string): string => {
        const clean = text.trim();
        if (!clean && text === '') return `<c r="${ref}" s="${style}"/>`;
        const idx = internStr(text);
        return `<c r="${ref}" s="${style}" t="s"><v>${idx}</v></c>`;
      };

      const renderDateCell = (ref: string, style: number, iso: string): string => {
        const serial = isoToExcelSerial(iso);
        if (serial === null) return `<c r="${ref}" s="${style}"/>`;
        return `<c r="${ref}" s="${style}"><v>${serial}</v></c>`;
      };

      const renderNumCell = (ref: string, style: number, val: number): string => {
        if (!Number.isFinite(val) || val <= 0) return `<c r="${ref}" s="${style}"/>`;
        return `<c r="${ref}" s="${style}"><v>${round2(val)}</v></c>`;
      };

      const buildDataRow = (p: Pedido | null, r: number, esCompacta: boolean): string => {
        const odd = esCompacta ? true : r % 2 === 1;
        const ht = esCompacta ? '14.25' : '24.0';
        const sA = odd ? 15 : 48;
        const sB = odd ? 16 : 42;
        const sC = odd ? 17 : 43;
        const sD = odd ? 18 : 39;
        const sE = odd ? 18 : 39;
        const sF = odd ? 16 : 42;
        const sG = odd ? 19 : 45;
        const sH = odd ? 19 : 45;
        const sI = 21;
        const sJ = 22;
        const sK = odd ? 16 : 42;
        const sL = odd ? 16 : 42;
        const sM = odd ? 18 : 39;
        const sN = odd ? 18 : 39;
        const sO = odd ? 16 : 42;
        const sP = odd ? 17 : 43;
        const sQ = odd ? 25 : 44;

        const fI = `IF(OR(B${r}="",G${r}=""),"",G${r}-IF(ISNUMBER(H${r}),H${r},0))`;
        const fJ = `IF(B${r}="","",IF(G${r}="","Falta total",IF(I${r}<0,"A favor",IF(I${r}=0,"Pagado",IF(H${r}>0,"Pago parcial","Pendiente")))))`;

        if (!p) {
          return `<row r="${r}" ht="${ht}" customHeight="1"><c r="A${r}" s="${sA}"/><c r="B${r}" s="${sB}"/><c r="C${r}" s="${sC}"/><c r="D${r}" s="${sD}"/><c r="E${r}" s="${sE}"/><c r="F${r}" s="${sF}"/><c r="G${r}" s="${sG}"/><c r="H${r}" s="${sH}"/><c r="I${r}" s="${sI}" t="str"><f>${escapeXml(fI)}</f><v></v></c><c r="J${r}" s="${sJ}" t="str"><f>${escapeXml(fJ)}</f><v></v></c><c r="K${r}" s="${sK}"/><c r="L${r}" s="${sL}"/><c r="M${r}" s="${sM}"/><c r="N${r}" s="${sN}"/><c r="O${r}" s="${sO}"/><c r="P${r}" s="${sP}"/><c r="Q${r}" s="${sQ}"/></row>`;
        }

        const contacto = p.whatsapp || p.vendedor || '';
        const detalle = p.detalleResumen && p.detalleResumen !== '—' ? p.detalleResumen : '';
        const estadoPagoTxt =
          p.totalFacturado <= 0
            ? 'Falta total'
            : p.saldoPendiente <= 0.01
              ? 'Pagado'
              : p.montoAbonado > 0
                ? 'Pago parcial'
                : 'Pendiente';

        const cellI =
          p.totalFacturado > 0
            ? `<c r="I${r}" s="${sI}"><f>${escapeXml(fI)}</f><v>${round2(p.saldoPendiente)}</v></c>`
            : `<c r="I${r}" s="${sI}" t="str"><f>${escapeXml(fI)}</f><v></v></c>`;

        const cellJ =
          p.medioPago === 'Cheque' && !p.chequeAcreditado
            ? renderStrCell(
                `J${r}`,
                34,
                p.detalleCheque ||
                  (p.fechaAcreditacionCheque ? `Cheque (${formatShortDate(p.fechaAcreditacionCheque)})` : 'Cheque')
              )
            : `<c r="J${r}" s="${sJ}" t="str"><f>${escapeXml(fJ)}</f><v>${escapeXml(estadoPagoTxt)}</v></c>`;

        const armadoTxt =
          p.estadoArmado === 'Por armar' && p.totalFacturado === 0 && p.lineas.length === 0 ? '' : p.estadoArmado;
        const envioTxt = p.estadoEnvio === 'Pendiente' && !p.transporte && !p.nroSeguimiento ? '' : p.estadoEnvio;

        return `<row r="${r}" ht="${ht}" customHeight="1">${renderStrCell(`A${r}`, odd ? 27 : 35, p.numeroPedido)}${renderStrCell(`B${r}`, odd ? 28 : 36, p.cliente)}${renderStrCell(`C${r}`, odd ? 29 : 37, contacto)}${renderDateCell(`D${r}`, odd ? 30 : 38, p.fechaPedido)}${renderDateCell(`E${r}`, sE, p.despachoPrevisto)}${renderStrCell(`F${r}`, sF, detalle)}${renderNumCell(`G${r}`, sG, p.totalFacturado)}${renderNumCell(`H${r}`, sH, p.montoAbonado)}${cellI}${cellJ}${renderStrCell(`K${r}`, sK, armadoTxt)}${renderStrCell(`L${r}`, sL, envioTxt)}${renderDateCell(`M${r}`, sM, p.fechaDespacho ?? '')}${renderDateCell(`N${r}`, odd ? 30 : 38, p.fechaEntrega ?? '')}${renderStrCell(`O${r}`, sO, p.transporte)}${renderStrCell(`P${r}`, sP, p.nroSeguimiento)}${renderStrCell(`Q${r}`, sQ, p.observaciones)}</row>`;
      };

      const cantPedidos = pedidosOrdenados.length;
      const porArmarCount = pedidosOrdenados.filter((p) => p.estadoArmado !== 'Listo').length;
      const porDespacharCount = pedidosOrdenados.filter(
        (p) => p.estadoEnvio !== 'Despachado' && p.estadoEnvio !== 'Entregado'
      ).length;
      const saldoPorCobrarSum = round2(
        pedidosOrdenados.reduce((acc, p) => acc + (p.saldoPendiente > 0 ? p.saldoPendiente : 0), 0)
      );
      const atrasadosCount = pedidosOrdenados.filter((p) => esDespachoAtrasado(p)).length;

      let s1Xml = await s1File.async('string');
      s1Xml = s1Xml.replace(/(<c r="A6"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${cantPedidos}</v>`);
      s1Xml = s1Xml.replace(/(<c r="D6"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${porArmarCount}</v>`);
      s1Xml = s1Xml.replace(/(<c r="G6"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${porDespacharCount}</v>`);
      s1Xml = s1Xml.replace(/(<c r="J6"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${saldoPorCobrarSum}</v>`);
      s1Xml = s1Xml.replace(/(<c r="M6"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${atrasadosCount}</v>`);

      const maxS1 = Math.max(299, 10 + pedidosOrdenados.length);
      const rowsS1: string[] = [];
      for (let r = 11; r <= maxS1; r++) {
        const p = pedidosOrdenados[r - 11] ?? null;
        rowsS1.push(buildDataRow(p, r, false));
      }
      const s1Start = s1Xml.indexOf('<row r="11"');
      const s1EndMarker = maxS1 > 299 ? '</sheetData>' : '<row r="300"';
      const s1End = s1Xml.indexOf(s1EndMarker);
      if (s1Start !== -1 && s1End !== -1) {
        s1Xml = s1Xml.slice(0, s1Start) + rowsS1.join('') + s1Xml.slice(s1End);
      }
      zip.file('xl/worksheets/sheet1.xml', s1Xml);

      let s2Xml = await s2File.async('string');
      s2Xml = s2Xml.replace(/(<c r="F1"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${cantPedidos}</v>`);
      s2Xml = s2Xml.replace(/(<c r="H1"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${porArmarCount}</v>`);
      s2Xml = s2Xml.replace(/(<c r="J1"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${porDespacharCount}</v>`);
      s2Xml = s2Xml.replace(/(<c r="L1"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${saldoPorCobrarSum}</v>`);
      s2Xml = s2Xml.replace(/(<c r="N1"[^>]*><f>[\s\S]*?<\/f>)<v>[^<]*<\/v>/, `$1<v>${atrasadosCount}</v>`);

      const maxS2 = Math.max(303, 3 + pedidosOrdenados.length);
      const rowsS2: string[] = [];
      for (let r = 4; r <= maxS2; r++) {
        const p = pedidosOrdenados[r - 4] ?? null;
        rowsS2.push(buildDataRow(p, r, true));
      }
      const s2Start = s2Xml.indexOf('<row r="4"');
      const s2EndMarker = maxS2 > 303 ? '</sheetData>' : '<row r="304"';
      const s2End = s2Xml.indexOf(s2EndMarker);
      if (s2Start !== -1 && s2End !== -1) {
        s2Xml = s2Xml.slice(0, s2Start) + rowsS2.join('') + s2Xml.slice(s2End);
      }
      zip.file('xl/worksheets/sheet2.xml', s2Xml);

      const newSsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${siList.length}" uniqueCount="${siList.length}">${siList.join('')}</sst>`;
      zip.file('xl/sharedStrings.xml', newSsXml);

      const outBytes = await zip.generateAsync({ type: 'uint8array' });
      return { buffer: outBytes, filename: nombreArchivo };
    }
  } catch {
    // Fallback a SheetJS puro si no está disponible la plantilla
  }

  const headers17: CellPrimitive[] = [
    'N.º pedido',
    'Cliente / negocio',
    'WhatsApp',
    'Fecha pedido',
    'Despacho previsto',
    'Detalle del pedido',
    'Total pedido',
    'Monto abonado',
    'Saldo pendiente',
    'Estado de pago',
    'Armado',
    'Envío',
    'Fecha despacho',
    'Fecha entrega',
    'Transporte',
    'N.º seguimiento',
    'Observaciones'
  ];

  const filas: CellPrimitive[][] = pedidosOrdenados.map((p) => [
    p.numeroPedido,
    p.cliente,
    p.whatsapp || p.vendedor || '',
    isoToExcelSerial(p.fechaPedido) ?? p.fechaPedido,
    isoToExcelSerial(p.despachoPrevisto) ?? (p.despachoPrevisto || ''),
    p.detalleResumen === '—' ? '' : p.detalleResumen,
    p.totalFacturado > 0 ? p.totalFacturado : null,
    p.montoAbonado > 0 ? p.montoAbonado : null,
    p.totalFacturado > 0 ? p.saldoPendiente : null,
    p.medioPago === 'Cheque' && !p.chequeAcreditado
      ? p.detalleCheque || `Cheque (${formatShortDate(p.fechaAcreditacionCheque)})`
      : p.totalFacturado <= 0
        ? 'Falta total'
        : p.estadoPago,
    p.estadoArmado,
    p.estadoEnvio,
    isoToExcelSerial(p.fechaDespacho ?? '') ?? (p.fechaDespacho || ''),
    isoToExcelSerial(p.fechaEntrega ?? '') ?? (p.fechaEntrega || ''),
    p.transporte || '',
    p.nroSeguimiento || '',
    p.observaciones || ''
  ]);

  const data: CellPrimitive[][] = [
    ['PEDIDOS MAYORISTA MIBI'],
    ['Una fila por pedido. Cargá desde la fila 11.'],
    [
      'Pedidos cargados',
      pedidosOrdenados.length,
      'Por armar',
      resumen.porArmar,
      'Cobrado en mano',
      resumen.cobradoEnMano,
      'En cheques 30d',
      resumen.enCheques,
      'Falta pago',
      resumen.faltaPago
    ],
    [],
    headers17,
    ...filas
  ];

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [
    { wch: 14 },
    { wch: 25 },
    { wch: 19 },
    { wch: 16 },
    { wch: 18 },
    { wch: 32 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 20 },
    { wch: 17 },
    { wch: 18 },
    { wch: 17 },
    { wch: 17 },
    { wch: 23 },
    { wch: 24 },
    { wch: 38 }
  ];
  aplicarFormatoMoneda(ws, [6, 7, 8]);

  const ref = ws['!ref'];
  if (ref) {
    const range = XLSX.utils.decode_range(ref);
    for (let r = 5; r <= range.e.r; r++) {
      for (const c of [3, 4, 12, 13]) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr] as SheetJSCellObject | undefined;
        if (cell && typeof cell.v === 'number') {
          cell.t = 'n';
          cell.z = 'dd/mm/yyyy';
        }
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'PEDIDOS MAYORISTA MIBI');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Uint8Array;
  return { buffer: new Uint8Array(buffer), filename: nombreArchivo };
};
