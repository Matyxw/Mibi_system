import JSZip from 'jszip';
import {
  analizarFilasTabulares,
  auditoriaVacia,
  construirAuditoriaDesdeTabular,
  extraerDatosCabeceraExcel,
  limpiarTextoInvisible,
  normalizarTexto
} from '@/lib/excel-parser';
import type {
  CellPrimitive,
  DatosClienteDetectados,
  ProductoCatalogo,
  ResultadoAnalisisArchivo
} from '@/lib/types';
import { normalizarCuit, validarCuit } from './numeros';

export type BloqueDocx =
  | { tipo: 'PARRAFO'; texto: string }
  | { tipo: 'TABLA'; filas: string[][] };

const decodeXmlEntities = (s: string): string =>
  s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(parseInt(dec, 10)));

const extractParagraphText = (xmlSegment: string): string => {
  const prepared = xmlSegment.replace(/<w:tab\b[^>]*\/>/g, ' ').replace(/<w:br\b[^>]*\/>/g, ' ');
  const parts: string[] = [];
  for (const m of prepared.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)) {
    if (m[1] !== undefined) parts.push(decodeXmlEntities(m[1]));
  }
  return limpiarTextoInvisible(parts.join(''));
};

export const extraerContenidoDocx = async (bytes: Uint8Array): Promise<BloqueDocx[]> => {
  const zip = await JSZip.loadAsync(bytes);
  const docFile =
    zip.file('word/document.xml') ??
    Object.values(zip.files).find((f) => f.name.toLowerCase() === 'word/document.xml');
  if (!docFile) return [];

  const xml = await docFile.async('string');
  const bloques: BloqueDocx[] = [];

  for (const m of xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>|<w:p(?:\s+[^>]*)?>[\s\S]*?<\/w:p>/g)) {
    const seg = m[0];
    if (seg.startsWith('<w:tbl')) {
      const filas: string[][] = [];
      for (const rm of seg.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)) {
        const celdas: string[] = [];
        for (const cm of rm[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)) {
          const pTexts: string[] = [];
          for (const pm of cm[0].matchAll(/<w:p(?:\s+[^>]*)?>[\s\S]*?<\/w:p>/g)) {
            const pt = extractParagraphText(pm[0]);
            if (pt) pTexts.push(pt);
          }
          celdas.push(limpiarTextoInvisible(pTexts.join(' ')));
        }
        if (celdas.some(Boolean)) {
          filas.push(celdas);
        }
      }
      if (filas.length > 0) {
        bloques.push({ tipo: 'TABLA', filas });
      }
    } else {
      const texto = extractParagraphText(seg);
      if (texto) {
        bloques.push({ tipo: 'PARRAFO', texto });
      }
    }
  }

  return bloques;
};

const parsearOpcionesMarcadas = (valor: string): string => {
  const clean = limpiarTextoInvisible(valor);
  if (!clean) return '';

  // Soporta patrones reales observados: "☐* Responsable Inscripto", "x☐ Responsable Inscripto", "☐ *Transferencia ☐* Cuenta corriente"
  const partes = clean.split('☐');
  const seleccionadas: string[] = [];

  for (let i = 0; i < partes.length; i++) {
    const seg = partes[i] ?? '';
    const prev = partes[i - 1] ?? '';
    const tieneMarcaAntes = i > 0 && /(?:\*|[xX✓✔☒])\s*$/.test(prev);
    const tieneMarcaDespues = /^\s*(?:\*|[xX✓✔☒])/.test(seg);
    if (i > 0 && (tieneMarcaAntes || tieneMarcaDespues)) {
      const etiqueta = seg
        .replace(/^\s*(?:\*|[xX✓✔☒])\s*/, '')
        .replace(/\s*(?:\*|[xX✓✔☒])\s*$/, '')
        .replace(/\s*Plazo:\s*_*\s*d[ií]as\b/i, '')
        .trim();
      if (etiqueta) seleccionadas.push(etiqueta);
    }
  }

  if (seleccionadas.length > 0) {
    return seleccionadas.join(', ');
  }
  return clean.replace(/[☐]/g, '').replace(/\s+/g, ' ').trim();
};

const ETIQUETAS_ALTA_CONOCIDAS = [
  'nombre o razon social',
  'condicion frente al iva',
  'domicilio fiscal',
  'telefono de contacto',
  'e-mail para facturacion',
  'direccion de entrega',
  'dias y horarios de recepcion'
] as const;

export const analizarDocx = async (
  bytes: Uint8Array,
  ctx: {
    fileName: string;
    catalogoPorCodigo: ReadonlyMap<string, ProductoCatalogo>;
    catalogoUnificado: ReadonlyArray<ProductoCatalogo>;
  }
): Promise<ResultadoAnalisisArchivo> => {
  let bloques: BloqueDocx[];
  try {
    bloques = await extraerContenidoDocx(bytes);
  } catch {
    return {
      ok: false,
      errorFatal: 'No se pudo leer el documento Word (.docx).',
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'NO_COMPATIBLE',
      auditoria: auditoriaVacia('BLOQUEADO'),
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados: null,
      nombreArchivo: ctx.fileName,
      avisos: [],
      visionDisponible: false,
      requiereConfirmacionHumana: false
    };
  }

  const todoTexto = normalizarTexto(
    bloques
      .map((b) => (b.tipo === 'PARRAFO' ? b.texto : b.filas.flat().join(' ')))
      .join(' ')
  );

  const hitsAlta = ETIQUETAS_ALTA_CONOCIDAS.filter((et) => todoTexto.includes(et)).length;
  const esAltaCliente = todoTexto.includes('alta de clientes') || hitsAlta >= 3;

  if (esAltaCliente) {
    type SeccionAlta = 'FISCAL' | 'CONTACTO' | 'ENTREGA' | 'PAGOS' | 'REFERENCIAS' | 'OTRA';
    let seccion: SeccionAlta = 'FISCAL';
    let cliente = '';
    let cuit = '';
    let whatsapp = '';
    let email = '';
    let domicilio = '';
    let localidadFiscal = '';
    let domicilioEntrega = '';
    let localidadEntrega = '';
    let horarioRecepcion = '';
    let condicionIva = '';
    let condicionPago = '';
    let contacto = '';
    let transporte = '';
    let observacionesFicha = '';
    const avisosExtra: string[] = [];

    const detectarSeccion = (txt: string): SeccionAlta | null => {
      const n = normalizarTexto(txt);
      if (n.includes('datos fiscales')) return 'FISCAL';
      if (n.includes('datos de contacto')) return 'CONTACTO';
      if (n.includes('datos de entrega')) return 'ENTREGA';
      if (n.includes('facturacion y pagos')) return 'PAGOS';
      if (n.includes('referencias comerciales')) return 'REFERENCIAS';
      return null;
    };

    for (const bloque of bloques) {
      if (bloque.tipo === 'PARRAFO') {
        const nextSec = detectarSeccion(bloque.texto);
        if (nextSec) seccion = nextSec;
        continue;
      }

      for (const fila of bloque.filas) {
        if (fila.length === 1) {
          const nextSec = detectarSeccion(fila[0] ?? '');
          if (nextSec) {
            seccion = nextSec;
            continue;
          }
        }
        if (seccion === 'REFERENCIAS') continue;

        const labelRaw = fila[0] ?? '';
        const valRaw = fila.slice(1).join(' ').trim();
        const label = normalizarTexto(labelRaw);
        const val = limpiarTextoInvisible(valRaw);
        if (!label || !val || /^_+$/.test(val)) continue;

        if (/nombre o razon social/.test(label)) {
          if (!cliente) cliente = val;
        } else if (/^c\.?u\.?i\.?t\.?$/.test(label)) {
          const digits = normalizarCuit(val);
          if (validarCuit(val)) {
            cuit = digits;
          } else if (digits.length === 11) {
            cuit = digits;
            avisosExtra.push(`Revisá el CUIT de la ficha (${val}): el dígito verificador no coincide con AFIP.`);
          }
        } else if (label.includes('condicion frente al iva')) {
          condicionIva = parsearOpcionesMarcadas(val);
        } else if (label.includes('domicilio fiscal')) {
          domicilio = val;
        } else if (label.includes('localidad')) {
          if (seccion === 'ENTREGA') {
            localidadEntrega = val;
          } else {
            localidadFiscal = val;
          }
        } else if (label.includes('persona de contacto')) {
          if (!contacto) contacto = val;
        } else if (label.includes('telefono de contacto')) {
          if (!whatsapp) whatsapp = val;
        } else if (label.includes('e-mail') && !label.includes('facturacion')) {
          if (!email) email = val;
        } else if (label.includes('direccion de entrega')) {
          domicilioEntrega = val;
        } else if (label.includes('horarios de recepcion')) {
          horarioRecepcion = val;
        } else if (label.includes('contacto que recibe') && !whatsapp) {
          const telMatch = val.match(/\d[\d\s\-]{6,}/);
          if (telMatch) whatsapp = telMatch[0].trim();
        } else if (label.includes('condicion de pago')) {
          condicionPago = parsearOpcionesMarcadas(val);
        } else if (label.includes('transporte') || label.includes('expreso')) {
          transporte = val;
        } else if (label === 'observaciones') {
          observacionesFicha = val;
        }
      }
    }

    const domCompleto = [domicilio, localidadFiscal].filter(Boolean).join(', ');
    const entCompleto = [domicilioEntrega, localidadEntrega].filter(Boolean).join(', ');
    const obsPartes: string[] = [];
    if (cuit) obsPartes.push(`CUIT: ${cuit}`);
    if (transporte) obsPartes.push(`Transporte: ${transporte}`);
    if (entCompleto) obsPartes.push(`Entrega: ${entCompleto}`);
    if (horarioRecepcion) obsPartes.push(`Horarios: ${horarioRecepcion}`);
    if (condicionIva) obsPartes.push(`IVA: ${condicionIva}`);
    if (condicionPago) obsPartes.push(`Pago ficha: ${condicionPago}`);
    if (observacionesFicha) obsPartes.push(observacionesFicha);

    const datosDetectados: DatosClienteDetectados = {
      ...(cliente ? { cliente } : {}),
      ...(cuit ? { cuit } : {}),
      ...(whatsapp ? { whatsapp } : {}),
      ...(email ? { email } : {}),
      ...(transporte ? { transporte } : {}),
      ...(domCompleto ? { domicilio: domCompleto } : {}),
      ...(entCompleto ? { domicilioEntrega: entCompleto } : {}),
      ...(horarioRecepcion ? { horarioRecepcion } : {}),
      ...(condicionIva ? { condicionIva } : {}),
      ...(condicionPago ? { condicionPago } : {}),
      ...(contacto ? { contacto } : {}),
      ...(obsPartes.length > 0 ? { observacionesSugeridas: obsPartes.join(' · ') } : {}),
      origen: 'FICHA_ALTA'
    };

    return {
      ok: true,
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'ALTA_CLIENTE_DOCX',
      auditoria: auditoriaVacia('EXACTO'),
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados,
      nombreArchivo: ctx.fileName,
      avisos: [
        'Detectamos una ficha de Alta de Cliente. Completamos los datos del cliente automáticamente; ahora podés subir la planilla o PDF del pedido.',
        ...avisosExtra
      ],
      visionDisponible: false,
      requiereConfirmacionHumana: false
    };
  }

  // Buscar tablas con pedido en Word (.docx)
  const grid: CellPrimitive[][] = [];
  for (const bloque of bloques) {
    if (bloque.tipo !== 'TABLA') continue;
    for (const fila of bloque.filas) {
      grid.push(fila.map((c) => c));
    }
  }

  const tieneCabeceraPedido = grid.some((row) => {
    const norms = row.map((c) => (c === null ? '' : normalizarTexto(String(c))));
    const hasCod = norms.some((t) => t.includes('codigo'));
    const hasQty = norms.some((t) => t.includes('cant') || t === 't1');
    return hasCod && hasQty;
  });

  if (tieneCabeceraPedido) {
    const tab = analizarFilasTabulares(grid, 'Word', ctx);
    const auditoria = construirAuditoriaDesdeTabular(tab, []);
    const datosDetectados = extraerDatosCabeceraExcel(
      grid,
      tab.headerRowIdx,
      ctx.fileName,
      tab.descuentoDetectado
    );
    return {
      ok: true,
      lineas: tab.lineas,
      advertencias: tab.advertencias,
      avisosPrecio: tab.avisosPrecio,
      tipoArchivo: 'PEDIDO_DOCX',
      auditoria,
      correccionesAutomaticas: tab.correcciones,
      hojasDisponibles: [],
      datosDetectados,
      nombreArchivo: ctx.fileName,
      avisos: [],
      visionDisponible: false,
      requiereConfirmacionHumana: false
    };
  }

  return {
    ok: false,
    errorFatal: 'No encontramos una ficha de Alta de Cliente ni una tabla de pedido en este documento Word.',
    lineas: [],
    advertencias: [],
    avisosPrecio: [],
    tipoArchivo: 'NO_COMPATIBLE',
    auditoria: auditoriaVacia('BLOQUEADO'),
    correccionesAutomaticas: [],
    hojasDisponibles: [],
    datosDetectados: null,
    nombreArchivo: ctx.fileName,
    avisos: [],
    visionDisponible: false,
    requiereConfirmacionHumana: false
  };
};
