import { getDocumentProxy } from 'unpdf';
import {
  analizarFilasTabulares,
  auditoriaVacia,
  construirAuditoriaDesdeTabular,
  identificarProductoPorDescripcion,
  limpiarTextoInvisible,
  normalizarTexto,
  resolverLinea,
  sugerirProductosCercanos
} from '@/lib/excel-parser';
import { round2 } from '@/lib/finanzas';
import type {
  AdvertenciaExcel,
  AuditoriaLectura,
  AvisoPrecioExcel,
  CellPrimitive,
  CorreccionAutomatica,
  DatosClienteDetectados,
  LineaPedido,
  ProductoCatalogo,
  ResultadoAnalisisArchivo
} from '@/lib/types';
import {
  normalizarCuit,
  parsearNumeroDocumento,
  validarCuit,
  formatearCuit
} from './numeros';

export interface CeldaPdf {
  x: number;
  texto: string;
}

export interface FilaPdf {
  y: number;
  celdas: CeldaPdf[];
}

export interface PaginaPdf {
  numero: number;
  filas: FilaPdf[];
}

export const extraerPaginasPdf = async (bytes: Uint8Array): Promise<PaginaPdf[]> => {
  const pdf = await getDocumentProxy(bytes, { verbosity: 0 });
  const paginas: PaginaPdf[] = [];

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();
    const rawItems: Array<{ x: number; y: number; texto: string }> = [];

    for (const item of tc.items) {
      if (!('str' in item)) continue;
      const texto = limpiarTextoInvisible(item.str);
      if (!texto) continue;
      const t = item.transform as number[] | undefined;
      const x = Number(t?.[4] ?? 0);
      const y = Number(t?.[5] ?? 0);
      rawItems.push({ x, y, texto });
    }

    rawItems.sort((a, b) => b.y - a.y || a.x - b.x);
    const filas: FilaPdf[] = [];
    for (const it of rawItems) {
      const ultima = filas[filas.length - 1];
      if (ultima && Math.abs(ultima.y - it.y) <= 2.5) {
        ultima.celdas.push({ x: it.x, texto: it.texto });
      } else {
        filas.push({ y: it.y, celdas: [{ x: it.x, texto: it.texto }] });
      }
    }
    for (const f of filas) {
      f.celdas.sort((a, b) => a.x - b.x);
    }
    paginas.push({ numero: p, filas });
  }

  return paginas;
};

export const clasificarPdf = (
  paginas: ReadonlyArray<PaginaPdf>
): 'IMAGEN_O_PDF_ESCANEADO' | 'FACTURA_AFIP_PDF' | 'PEDIDO_PDF' | 'NO_COMPATIBLE' => {
  let chars = 0;
  for (const pag of paginas) {
    for (const fila of pag.filas) {
      for (const c of fila.celdas) {
        chars += c.texto.replace(/\s+/g, '').length;
      }
    }
  }
  if (chars < 20) return 'IMAGEN_O_PDF_ESCANEADO';

  const muestra = normalizarTexto(
    paginas
      .flatMap((p) => p.filas.flatMap((f) => f.celdas.map((c) => c.texto)))
      .join(' ')
  );

  if (
    (muestra.includes('factura') ||
      muestra.includes('nota de credito') ||
      muestra.includes('nota de debito') ||
      muestra.includes('recibo')) &&
    (muestra.includes('comp. nro') || muestra.includes('cae')) &&
    (muestra.includes('importe total') || muestra.includes('importe neto gravado'))
  ) {
    return 'FACTURA_AFIP_PDF';
  }

  const tieneHeaderCodigo = paginas.some((p) =>
    p.filas.some((f) => f.celdas.some((c) => /^codigos?$/i.test(normalizarTexto(c.texto))))
  );
  const tieneCodigos6xxx = paginas.some((p) =>
    p.filas.some((f) => f.celdas.some((c) => /\b6\d{3}\b/.test(c.texto)))
  );

  if (tieneHeaderCodigo && tieneCodigos6xxx) {
    return 'PEDIDO_PDF';
  }

  return 'NO_COMPATIBLE';
};

const seleccionarPaginasOriginal = (paginas: ReadonlyArray<PaginaPdf>): PaginaPdf[] => {
  const conEtiquetaOriginal = paginas.filter((p) =>
    p.filas.some((f) => f.celdas.some((c) => /^ORIGINAL$/i.test(c.texto.trim())))
  );
  if (conEtiquetaOriginal.length > 0) {
    return conEtiquetaOriginal;
  }

  const tieneAlgunaEtiquetaCopia = paginas.some((p) =>
    p.filas.some((f) => f.celdas.some((c) => /^(?:DUPLICADO|TRIPLICADO)$/i.test(c.texto.trim())))
  );
  if (!tieneAlgunaEtiquetaCopia) {
    for (const p of paginas) {
      for (const f of p.filas) {
        for (const c of f.celdas) {
          const m = c.texto.match(/^P[áa]g\.\s*\d+\s*\/\s*(\d+)$/i);
          if (m && m[1]) {
            const totalOriginal = parseInt(m[1], 10);
            if (totalOriginal > 0 && totalOriginal <= paginas.length) {
              return paginas.slice(0, totalOriginal);
            }
          }
        }
      }
    }
  }
  return [...paginas];
};

const dmyToIso = (dmy: string): string | undefined => {
  const m = dmy.trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return undefined;
  return `${m[3]}-${m[2]}-${m[1]}`;
};

interface ItemFacturaPdfRaw {
  pagina: number;
  fila: number;
  codigo: string;
  descripcion: string;
  cantidad: number;
  precioUnitario: number;
  bonifPorc: number;
  subtotal: number;
}

const deducirProductoYVarianteFactura = (
  descripcion: string,
  catalogoUnificado: ReadonlyArray<ProductoCatalogo>
): { producto: ProductoCatalogo; codigoVariante: string } | null => {
  const prod = identificarProductoPorDescripcion(descripcion, catalogoUnificado);
  if (!prod) return null;
  if (prod.variantes.length === 1) {
    return { producto: prod, codigoVariante: prod.variantes[0]!.codigo };
  }
  const descNorm = normalizarTexto(descripcion);
  const matchedVars = prod.variantes.filter((v) =>
    v.aliases.some((al) => {
      const escaped = al.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
      return new RegExp(`\\b${escaped}\\b`, 'i').test(descNorm);
    })
  );
  if (matchedVars.length === 1) {
    return { producto: prod, codigoVariante: matchedVars[0]!.codigo };
  }
  return null;
};

export const parsearFacturaAfipDesdePaginas = (
  paginasInput: ReadonlyArray<PaginaPdf>,
  ctx: {
    fileName: string;
    catalogoPorCodigo: ReadonlyMap<string, ProductoCatalogo>;
    catalogoUnificado: ReadonlyArray<ProductoCatalogo>;
  }
): ResultadoAnalisisArchivo => {
  const paginas = seleccionarPaginasOriginal(paginasInput);
  const cuitEmisorEsperado =
    (typeof process !== 'undefined' ? process.env.MIBI_CUIT_EMISOR?.replace(/\D/g, '') : '') ||
    '27310986599';

  let tipoDocNoPedido: 'NOTA_CREDITO' | 'NOTA_DEBITO' | 'RECIBO' | null = null;
  let cuitEmisorDetectado = '';

  let cliente = '';
  let cuit = '';
  let domicilio = '';
  let numeroComprobante = '';
  let tipoComprobante = 'Factura A';
  let fechaPedido = '';

  let netoDeclarado: number | null = null;
  let subtotalFacturaB: number | null = null;
  let iva21Declarado: number | null = null;
  let ivaOtrosSuma = 0;
  let otrosTributosDeclarado = 0;
  let noGravadoDeclarado = 0;
  let totalFacturaDeclarado: number | null = null;

  const items: ItemFacturaPdfRaw[] = [];
  const itemsNegativos: ItemFacturaPdfRaw[] = [];
  const avisos: string[] = [];

  for (let pIdx = 0; pIdx < paginas.length; pIdx++) {
    const pag = paginas[pIdx]!;
    const filas = pag.filas;

    // Detectar header de tabla de ítems en esta página
    let colHeaders: Array<{ key: string; x: number }> = [
      { key: 'codigo', x: 19 },
      { key: 'producto', x: 59 },
      { key: 'cantidad', x: 239 },
      { key: 'umedida', x: 286 },
      { key: 'precio', x: 335 },
      { key: 'bonif', x: 388 },
      { key: 'subtotal', x: 435 },
      { key: 'iva', x: 484 },
      { key: 'subIva', x: 524 }
    ];
    let itemsTopY = 568;
    let itemsBottomY = 160;

    for (const f of filas) {
      const joined = f.celdas.map((c) => normalizarTexto(c.texto)).join(' ');
      if (joined.includes('codigo') && joined.includes('producto') && joined.includes('cantidad')) {
        itemsTopY = f.y;
        const detected: Array<{ key: string; x: number }> = [];
        for (const c of f.celdas) {
          const n = normalizarTexto(c.texto);
          if (n.includes('codigo')) detected.push({ key: 'codigo', x: c.x });
          else if (n.includes('producto')) detected.push({ key: 'producto', x: c.x });
          else if (n.includes('cantidad')) detected.push({ key: 'cantidad', x: c.x });
          else if (n.includes('medida')) detected.push({ key: 'umedida', x: c.x });
          else if (n.includes('precio')) detected.push({ key: 'precio', x: c.x });
          else if (n.includes('bonif')) detected.push({ key: 'bonif', x: c.x });
          else if (n === 'subtotal') detected.push({ key: 'subtotal', x: c.x });
          else if (n.includes('alicuota')) detected.push({ key: 'iva', x: c.x });
          else if (n.includes('c/iva')) detected.push({ key: 'subIva', x: c.x });
        }
        if (detected.length >= 6) {
          if (!detected.some((d) => d.key === 'iva')) {
            const subH = detected.find((d) => d.key === 'subtotal');
            const subIvaH = detected.find((d) => d.key === 'subIva');
            if (subH && subIvaH && subIvaH.x - subH.x > 60) {
              detected.push({ key: 'iva', x: Math.round((subH.x + subIvaH.x) / 2) });
            } else {
              detected.push({ key: 'iva', x: 484 });
            }
          }
          detected.sort((a, b) => a.x - b.x);
          colHeaders = detected;
        }
      }
      if (f.celdas.some((c) => /^Importe (?:Neto Gravado|Neto No Gravado|Op\. Exentas|Otros Tributos|Total):/i.test(c.texto))) {
        itemsBottomY = Math.max(itemsBottomY, f.y + 5);
      }
    }

    const headerYRef = colHeaders;
    const assignCol = (x: number): string => {
      let best = headerYRef[0]?.key ?? 'codigo';
      for (const h of headerYRef) {
        if (h.x <= x + 6) {
          best = h.key;
        }
      }
      return best;
    };

    const tieneImporteTotalEnPagina = filas.some((f) =>
      f.celdas.some((c) => /^Importe Total:\s*\$/i.test(c.texto))
    );

    for (let rIdx = 0; rIdx < filas.length; rIdx++) {
      const fila = filas[rIdx]!;
      const textos = fila.celdas.map((c) => c.texto);
      const filaNorm = normalizarTexto(textos.join(' '));

      // Bloque superior: tipo de comprobante, COD. XXX y CUIT emisor
      if (fila.y > 670) {
        if (/nota\s+de\s+credito/i.test(filaNorm)) tipoDocNoPedido = 'NOTA_CREDITO';
        else if (/nota\s+de\s+debito/i.test(filaNorm)) tipoDocNoPedido = 'NOTA_DEBITO';
        else if (fila.y > 740 && /\brecibo\b/i.test(filaNorm)) tipoDocNoPedido = 'RECIBO';

        for (const c of fila.celdas) {
          if (c.x >= 260 && c.x <= 320 && /^[ABC]$/.test(c.texto)) {
            tipoComprobante = `Factura ${c.texto}`;
          }
          if (/COD\.\s*0*(?:03|08|13)\b/i.test(c.texto)) tipoDocNoPedido = 'NOTA_CREDITO';
          else if (/COD\.\s*0*(?:02|07|12)\b/i.test(c.texto)) tipoDocNoPedido = 'NOTA_DEBITO';
          else if (/COD\.\s*0*1\b/i.test(c.texto)) tipoComprobante = 'Factura A';
          else if (/COD\.\s*0*6\b/i.test(c.texto)) tipoComprobante = 'Factura B';
          else if (/COD\.\s*11\b/i.test(c.texto)) tipoComprobante = 'Factura C';
        }

        const esReceptorRow = fila.celdas.some((c) =>
          /apellido y nombre|raz[óo]n social/i.test(c.texto) && c.x > 180
        );
        if (!cuitEmisorDetectado && fila.y > 675 && !esReceptorRow) {
          for (const c of fila.celdas) {
            const inlineMatch = c.texto.match(/c\.?u\.?i\.?t\.?\s*:?\s*(\d{2}-?\d{8}-?\d)/i);
            if (inlineMatch?.[1]) {
              const d = normalizarCuit(inlineMatch[1]);
              if (validarCuit(d)) {
                cuitEmisorDetectado = d;
                break;
              }
            }
          }
          if (!cuitEmisorDetectado) {
            const cuitLabel = fila.celdas.find((c) => /^c\.?u\.?i\.?t\.?:?$/i.test(c.texto.trim()));
            if (cuitLabel) {
              const cuitEmCell = fila.celdas.find((c) => {
                if (c.x <= cuitLabel.x) return false;
                const d = normalizarCuit(c.texto);
                return d.length === 11;
              });
              if (cuitEmCell) {
                const d = normalizarCuit(cuitEmCell.texto);
                if (validarCuit(d)) cuitEmisorDetectado = d;
              }
            }
          }
        }
      }

      // Punto de venta + número de comprobante
      if (!numeroComprobante) {
        const pv = fila.celdas.find((c) => c.x > 350 && /^\d{5}$/.test(c.texto));
        const nro = fila.celdas.find((c) => c.x > 450 && /^\d{8}$/.test(c.texto));
        if (pv && nro) {
          numeroComprobante = `${pv.texto}-${nro.texto}`;
        }
      }

      // Fecha de Emisión
      if (!fechaPedido && fila.y >= 720 && fila.y <= 745) {
        const fCell = fila.celdas.find((c) => c.x > 380 && /^\d{2}\/\d{2}\/\d{4}$/.test(c.texto));
        if (fCell) {
          fechaPedido = dmyToIso(fCell.texto) ?? '';
        }
      }

      // Receptor (Razón social + CUIT)
      const idxRazonReceptor = fila.celdas.findIndex(
        (c) => /apellido y nombre|raz[óo]n social/i.test(c.texto) && c.x > 180
      );
      if (!cliente && idxRazonReceptor !== -1) {
        const labelX = fila.celdas[idxRazonReceptor]!.x;
        const nombreCell = fila.celdas.find((c) => c.x > labelX + 10);
        if (nombreCell && !/GOGNI MELISA/i.test(nombreCell.texto)) {
          cliente = nombreCell.texto;
        }
        const cuitCell = fila.celdas.find((c) => {
          const d = normalizarCuit(c.texto);
          return d.length === 11 && d !== cuitEmisorEsperado;
        });
        if (cuitCell && !cuit) {
          const d = normalizarCuit(cuitCell.texto);
          if (validarCuit(d)) cuit = d;
        }
      }

      // Domicilio comercial del receptor (el de la derecha, x > 240)
      if (!domicilio) {
        const domLabel = fila.celdas.find((c) => /domicilio comercial:/i.test(c.texto) && c.x > 200);
        if (domLabel) {
          const domVal = fila.celdas.find((c) => c.x > domLabel.x + 20);
          if (domVal) domicilio = domVal.texto;
        }
      }

      // Filas de arrastre multi-página que se deben ignorar
      if (
        fila.celdas.some(
          (c) =>
            /^TRANSPORTE:/i.test(c.texto) ||
            /^Subtotal\s*\(incluye transporte\):/i.test(c.texto) ||
            (/^Subtotal:$/i.test(c.texto) && !tieneImporteTotalEnPagina)
        )
      ) {
        continue;
      }

      // Bloque de totales al pie
      if (fila.celdas.some((c) => /^Importe Neto Gravado:/i.test(c.texto))) {
        const val = parsearNumeroDocumento(fila.celdas[fila.celdas.length - 1]?.texto);
        if (val > 0) netoDeclarado = val;
        continue;
      }
      if (fila.celdas.some((c) => /^Importe Neto No Gravado:|^Importe Op\. Exentas:/i.test(c.texto))) {
        const val = parsearNumeroDocumento(fila.celdas[fila.celdas.length - 1]?.texto);
        if (val > 0) noGravadoDeclarado = round2(noGravadoDeclarado + val);
        continue;
      }
      if (fila.celdas.some((c) => /^Subtotal:\s*\$?/i.test(c.texto)) && tieneImporteTotalEnPagina) {
        const val = parsearNumeroDocumento(fila.celdas[fila.celdas.length - 1]?.texto);
        if (val > 0) subtotalFacturaB = val;
        continue;
      }
      if (fila.celdas.some((c) => /^IVA 21%:/i.test(c.texto))) {
        iva21Declarado = parsearNumeroDocumento(fila.celdas[fila.celdas.length - 1]?.texto);
        continue;
      }
      if (fila.celdas.some((c) => /^IVA (?:27|10\.5|5|2\.5)%:/i.test(c.texto))) {
        ivaOtrosSuma = round2(
          ivaOtrosSuma + parsearNumeroDocumento(fila.celdas[fila.celdas.length - 1]?.texto)
        );
        continue;
      }
      if (fila.celdas.some((c) => /^Importe Otros Tributos:/i.test(c.texto) && c.x >= 340)) {
        otrosTributosDeclarado = parsearNumeroDocumento(fila.celdas[fila.celdas.length - 1]?.texto);
        continue;
      }
      if (fila.celdas.some((c) => /^Importe Total:/i.test(c.texto))) {
        const val = parsearNumeroDocumento(fila.celdas[fila.celdas.length - 1]?.texto);
        if (val > 0) totalFacturaDeclarado = val;
        continue;
      }

      // Ítems dentro de la zona de tabla
      if (fila.y >= itemsTopY - 2 || fila.y <= itemsBottomY) continue;

      const tieneUnidades = fila.celdas.some((c) => /^unidades$/i.test(c.texto));
      if (tieneUnidades) {
        let codigo = '';
        const descParts: string[] = [];
        let cantidad = 0;
        let precioUnitario = 0;
        let bonifPorc = 0;
        let subtotal = 0;

        for (const c of fila.celdas) {
          const col = assignCol(c.x);
          if (col === 'codigo') {
            codigo = c.texto.trim();
          } else if (col === 'producto') {
            descParts.push(c.texto);
          } else if (col === 'cantidad') {
            cantidad = parsearNumeroDocumento(c.texto);
          } else if (col === 'precio') {
            precioUnitario = parsearNumeroDocumento(c.texto);
          } else if (col === 'bonif') {
            bonifPorc = parsearNumeroDocumento(c.texto);
          } else if (col === 'subtotal') {
            subtotal = parsearNumeroDocumento(c.texto);
          }
        }

        const itemObj: ItemFacturaPdfRaw = {
          pagina: pag.numero,
          fila: rIdx + 1,
          codigo,
          descripcion: limpiarTextoInvisible(descParts.join(' ')),
          cantidad,
          precioUnitario,
          bonifPorc,
          subtotal
        };
        if (cantidad > 0 && subtotal >= 0) {
          items.push(itemObj);
        } else {
          itemsNegativos.push(itemObj);
        }
      } else if (
        items.length > 0 &&
        !/^(?:alicuota|iva|pag\.|cae|comprobante|esta agencia)/i.test(filaNorm) &&
        fila.celdas.every((c) => assignCol(c.x) === 'producto')
      ) {
        const prev = items[items.length - 1]!;
        if (prev.pagina === pag.numero) {
          prev.descripcion = limpiarTextoInvisible(`${prev.descripcion} ${textos.join(' ')}`);
        }
      }
    }
  }

  // Bloqueos duros de cumplimiento AFIP / ARCA
  if (tipoDocNoPedido) {
    const msgMap: Record<typeof tipoDocNoPedido, string> = {
      NOTA_CREDITO:
        'Este archivo es una Nota de Crédito AFIP (devolución o anulación), no un pedido de compra.',
      NOTA_DEBITO: 'Este archivo es una Nota de Débito AFIP, no un pedido de compra.',
      RECIBO: 'Este archivo es un Recibo AFIP, no un pedido de compra.'
    };
    return {
      ok: false,
      errorFatal: msgMap[tipoDocNoPedido],
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'COMPROBANTE_NO_PEDIDO',
      auditoria: auditoriaVacia('BLOQUEADO'),
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados: null,
      comentariosDetectados: [],
      nombreArchivo: ctx.fileName,
      avisos: [],
      visionDisponible: false,
      requiereConfirmacionHumana: false
    };
  }

  if (cuitEmisorDetectado && cuitEmisorDetectado !== cuitEmisorEsperado) {
    return {
      ok: false,
      errorFatal: `Esta factura AFIP fue emitida por CUIT ${formatearCuit(cuitEmisorDetectado)} y no corresponde a Mibi®.`,
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'COMPROBANTE_NO_PEDIDO',
      auditoria: auditoriaVacia('BLOQUEADO'),
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados: null,
      comentariosDetectados: [],
      nombreArchivo: ctx.fileName,
      avisos: [],
      visionDisponible: false,
      requiereConfirmacionHumana: false
    };
  }

  // Resolver ítems al catálogo unificado agrupando variantes del mismo grupo
  const mapaGrupos = new Map<string, { prod: ProductoCatalogo; vars: Map<string, number> }>();
  const advertencias: AdvertenciaExcel[] = [];
  const avisosPrecioMap = new Map<string, AvisoPrecioExcel>();
  const correcciones: CorreccionAutomatica[] = [];
  const bonificacionesUsadas = new Set<number>();
  let filasOmitidasConDatos = 0;

  for (const neg of itemsNegativos) {
    filasOmitidasConDatos += 1;
    advertencias.push({
      hoja: `Pág. ${neg.pagina}`,
      fila: neg.fila,
      codigo: neg.codigo,
      mensaje: `Fila con cantidad o subtotal negativo/inválido en factura (${neg.cantidad} u. · $${neg.subtotal}).`,
      id: `Pág. ${neg.pagina}#${neg.fila}#${neg.codigo || 'NEG'}#CANTIDAD_INVALIDA`,
      tipo: 'CANTIDAD_INVALIDA',
      codigosDetectados: neg.codigo ? [neg.codigo] : [],
      descripcionExcel: neg.descripcion,
      precioExcelReferencia: neg.precioUnitario !== 0 ? Math.abs(neg.precioUnitario) : null,
      cantidadDetectada: null,
      cantidadRaw: String(neg.cantidad),
      notaDetectada: neg.descripcion,
      productoExistenteIdGrupo: neg.codigo ? (ctx.catalogoPorCodigo.get(neg.codigo)?.idGrupo ?? null) : null,
      sugerenciasCatalogo: []
    });
  }

  const esFacturaB = tipoComprobante === 'Factura B';
  const esFacturaC = tipoComprobante === 'Factura C';

  for (const it of items) {
    bonificacionesUsadas.add(round2(it.bonifPorc));

    let prod: ProductoCatalogo | undefined;
    let codigoEfectivo = it.codigo.trim().toUpperCase();

    if (codigoEfectivo) {
      prod = ctx.catalogoPorCodigo.get(codigoEfectivo);
    }

    if (!prod && !codigoEfectivo) {
      const deducido = deducirProductoYVarianteFactura(it.descripcion, ctx.catalogoUnificado);
      if (deducido) {
        prod = deducido.producto;
        codigoEfectivo = deducido.codigoVariante;
        correcciones.push({
          id: `AFIP#P${it.pagina}F${it.fila}#deducido->${codigoEfectivo}`,
          hoja: `Pág. ${it.pagina}`,
          fila: it.fila,
          original: 'Sin código',
          corregidoA: codigoEfectivo,
          productoNombre: prod.nombre,
          motivo: 'Código deducido por la descripción de la factura',
          lineaId: prod.idGrupo
        });
      } else {
        filasOmitidasConDatos += 1;
        advertencias.push({
          hoja: `Pág. ${it.pagina}`,
          fila: it.fila,
          codigo: '',
          mensaje: `Fila de factura sin código (${it.cantidad} u.): "${it.descripcion}"`,
          id: `Pág. ${it.pagina}#${it.fila}#FILA_SIN_CODIGO`,
          tipo: 'FILA_SIN_CODIGO',
          codigosDetectados: [],
          descripcionExcel: it.descripcion,
          precioExcelReferencia: it.precioUnitario > 0 ? it.precioUnitario : null,
          cantidadDetectada: Number.isInteger(it.cantidad) ? it.cantidad : Math.round(it.cantidad),
          cantidadRaw: String(it.cantidad),
          notaDetectada: it.descripcion,
          productoExistenteIdGrupo: null,
          sugerenciasCatalogo: sugerirProductosCercanos('', it.descripcion, ctx.catalogoUnificado)
        });
        continue;
      }
    }

    if (!prod) {
      advertencias.push({
        hoja: `Pág. ${it.pagina}`,
        fila: it.fila,
        codigo: codigoEfectivo,
        mensaje: `Código desconocido en factura (${codigoEfectivo}) con cantidad ${it.cantidad}`,
        id: `Pág. ${it.pagina}#${it.fila}#${codigoEfectivo}#CODIGO_DESCONOCIDO`,
        tipo: 'CODIGO_DESCONOCIDO',
        codigosDetectados: codigoEfectivo ? [codigoEfectivo] : [],
        descripcionExcel: it.descripcion,
        precioExcelReferencia: it.precioUnitario > 0 ? it.precioUnitario : null,
        cantidadDetectada: Number.isInteger(it.cantidad) ? it.cantidad : Math.round(it.cantidad),
        cantidadRaw: String(it.cantidad),
        notaDetectada: it.descripcion,
        productoExistenteIdGrupo: null,
        sugerenciasCatalogo: sugerirProductosCercanos(codigoEfectivo, it.descripcion, ctx.catalogoUnificado)
      });
      continue;
    }

    const precioUnitarioNetoFactura = esFacturaB
      ? round2(it.precioUnitario / 1.21)
      : it.precioUnitario;

    if (
      precioUnitarioNetoFactura > 0 &&
      Math.abs(precioUnitarioNetoFactura - prod.precioNeto) > 1.5 &&
      !avisosPrecioMap.has(codigoEfectivo)
    ) {
      avisosPrecioMap.set(codigoEfectivo, {
        codigo: codigoEfectivo,
        nombreProducto: prod.nombre,
        precioPlanilla: precioUnitarioNetoFactura,
        precioVigente: prod.precioNeto
      });
    }

    let entry = mapaGrupos.get(prod.idGrupo);
    if (!entry) {
      entry = { prod, vars: new Map<string, number>() };
      mapaGrupos.set(prod.idGrupo, entry);
    }
    entry.vars.set(codigoEfectivo, (entry.vars.get(codigoEfectivo) ?? 0) + it.cantidad);
  }

  const lineas: LineaPedido[] = [];
  for (const { prod, vars } of mapaGrupos.values()) {
    const qtyGrupo = Array.from(vars.values()).reduce((s, q) => s + q, 0);
    let notaGrupo = '';
    if (prod.variantes.length >= 2) {
      notaGrupo = Array.from(vars.entries())
        .map(([cod, q]) => `${cod}: ${q}`)
        .join(' / ');
    }
    const linea = resolverLinea(prod, Math.round(qtyGrupo), notaGrupo);
    lineas.push(linea);
    for (const c of correcciones) {
      if (c.lineaId === prod.idGrupo) {
        c.lineaId = linea.idLinea;
      }
    }
  }

  const unidadesDeclaradasArchivo = round2(items.reduce((s, it) => s + it.cantidad, 0));
  const unidadesLeidas = lineas.reduce((s, l) => s + l.cantidadTotal, 0);
  const unidadesEnAlertas = advertencias.reduce((s, a) => s + (a.cantidadDetectada ?? 0), 0);
  const diferenciaUnidades = unidadesLeidas - unidadesDeclaradasArchivo;

  const subtotalNetoCalculado = round2(items.reduce((s, it) => s + it.subtotal, 0));
  const totalDeclaradoArchivo =
    esFacturaB
      ? subtotalFacturaB ?? totalFacturaDeclarado ?? null
      : esFacturaC
        ? netoDeclarado ?? subtotalFacturaB ?? totalFacturaDeclarado ?? null
        : netoDeclarado ?? subtotalFacturaB ?? null;
  const diferenciaImporte =
    totalDeclaradoArchivo !== null ? round2(subtotalNetoCalculado - totalDeclaradoArchivo) : null;

  const motivos: string[] = [];
  let estadoIntegridad: AuditoriaLectura['estadoIntegridad'] = 'EXACTO';

  if (diferenciaUnidades !== 0) {
    estadoIntegridad = 'BLOQUEADO';
    motivos.push(
      `La factura suma ${unidadesDeclaradasArchivo} unidades, pero se cargaron ${unidadesLeidas} (diferencia de ${diferenciaUnidades} u.).`
    );
  }
  if (diferenciaImporte !== null && Math.abs(diferenciaImporte) > 1) {
    estadoIntegridad = 'BLOQUEADO';
    motivos.push(
      `El subtotal de ítems ($${subtotalNetoCalculado.toFixed(2)}) no coincide con el ${esFacturaB ? 'Subtotal' : 'Neto Gravado'} de la factura ($${totalDeclaradoArchivo?.toFixed(2)}).`
    );
  }

  if (esFacturaB && totalFacturaDeclarado !== null) {
    const baseB = subtotalFacturaB ?? subtotalNetoCalculado;
    const esperadoB = round2(baseB + otrosTributosDeclarado);
    if (Math.abs(esperadoB - totalFacturaDeclarado) > 1) {
      estadoIntegridad = 'BLOQUEADO';
      motivos.push(
        `El total impreso en Factura B ($${totalFacturaDeclarado.toFixed(2)}) no coincide con Subtotal + Otros Tributos ($${esperadoB.toFixed(2)}).`
      );
    }
  } else if (esFacturaC && totalFacturaDeclarado !== null) {
    const esperadoC = round2(subtotalNetoCalculado + otrosTributosDeclarado);
    if (Math.abs(esperadoC - totalFacturaDeclarado) > 1) {
      estadoIntegridad = 'BLOQUEADO';
      motivos.push(
        `El total impreso en Factura C ($${totalFacturaDeclarado.toFixed(2)}) no coincide con la suma de ítems ($${esperadoC.toFixed(2)}).`
      );
    }
  } else if (
    totalFacturaDeclarado !== null &&
    netoDeclarado !== null &&
    iva21Declarado !== null
  ) {
    const esperadoConIva = round2(
      netoDeclarado + iva21Declarado + ivaOtrosSuma + otrosTributosDeclarado + noGravadoDeclarado
    );
    if (Math.abs(esperadoConIva - totalFacturaDeclarado) > 1) {
      estadoIntegridad = 'BLOQUEADO';
      motivos.push(
        `El total impreso en la factura ($${totalFacturaDeclarado.toFixed(2)}) no coincide con Neto + IVA + Otros Tributos ($${esperadoConIva.toFixed(2)}).`
      );
    }
  }

  if (otrosTributosDeclarado > 0) {
    avisos.push(
      `La factura incluye $${otrosTributosDeclarado.toFixed(2)} en percepciones / otros tributos.`
    );
  }
  if (esFacturaC) {
    avisos.push('Comprobante Factura C detectado.');
  }

  let descuentoPorc: number | undefined;
  const bonifsArray = Array.from(bonificacionesUsadas);
  if (bonifsArray.length > 1) {
    avisos.push('La factura tiene bonificaciones distintas por artículo; revisá el descuento.');
    if (estadoIntegridad === 'EXACTO') {
      estadoIntegridad = 'REVISAR';
      motivos.push('Bonificación mixta entre artículos de la factura.');
    }
  } else if (bonifsArray.length === 1 && (bonifsArray[0] ?? 0) > 0) {
    descuentoPorc = bonifsArray[0];
  }

  if (estadoIntegridad === 'EXACTO' && (advertencias.length > 0 || correcciones.length > 0)) {
    estadoIntegridad = 'REVISAR';
    if (advertencias.length > 0) {
      motivos.push(`Hay ${advertencias.length} fila(s) sin resolver en la factura.`);
    }
  }

  const refFactura = numeroComprobante ? `${tipoComprobante} AFIP ${numeroComprobante}` : 'Factura AFIP';
  const obsSugerida =
    otrosTributosDeclarado > 0
      ? `${refFactura} · Percepciones: $${otrosTributosDeclarado.toFixed(2)}`
      : refFactura;

  const datosDetectados: DatosClienteDetectados = {
    ...(cliente ? { cliente } : {}),
    ...(cuit ? { cuit } : {}),
    ...(domicilio ? { domicilio } : {}),
    ...(numeroComprobante ? { numeroComprobante } : {}),
    tipoComprobante,
    ...(fechaPedido ? { fechaPedido } : {}),
    ...(descuentoPorc !== undefined ? { descuentoPorc } : {}),
    observacionesSugeridas: cuit ? `${obsSugerida} · CUIT: ${cuit}` : obsSugerida,
    origen: 'FACTURA'
  };

  return {
    ok: true,
    lineas,
    advertencias,
    avisosPrecio: Array.from(avisosPrecioMap.values()),
    tipoArchivo: 'FACTURA_AFIP_PDF',
    auditoria: {
      unidadesLeidas,
      unidadesDeclaradasArchivo,
      diferenciaUnidades,
      unidadesEnAlertas,
      subtotalNetoCalculado,
      totalDeclaradoArchivo,
      diferenciaImporte,
      filasOmitidasConDatos,
      estadoIntegridad,
      motivos,
      ivaDeclaradoArchivo: iva21Declarado,
      totalFacturadoDeclaradoArchivo: totalFacturaDeclarado,
      duplicadosEnArchivo: []
    },
    correccionesAutomaticas: correcciones,
    hojasDisponibles: [],
    datosDetectados,
    comentariosDetectados: [],
    nombreArchivo: ctx.fileName,
    avisos,
    visionDisponible: false,
    requiereConfirmacionHumana: false
  };
};

export const parsearPedidoPdfDesdePaginas = (
  paginas: ReadonlyArray<PaginaPdf>,
  ctx: {
    fileName: string;
    catalogoPorCodigo: ReadonlyMap<string, ProductoCatalogo>;
    catalogoUnificado: ReadonlyArray<ProductoCatalogo>;
  }
): ResultadoAnalisisArchivo => {
  const grid: CellPrimitive[][] = [
    ['Codigos', 'Codigo 2', 'Producto', 'Precio', 'Cantidades Pedidas', 'Subtotal', 'Detalle Talles/ Colores']
  ];

  interface FilaPedidoRaw {
    y: number;
    codigos: string[];
    descParts: string[];
    precio: string;
    qty: string;
    subtotal: string;
    notas: string[];
  }

  let totalSinIva: number | null = null;

  for (const pag of paginas) {
    if (pag.filas.length === 0) continue;

    let precioX = 485;
    let cantX = 556;
    let subtotalX = 578;
    let notaX = 652;
    let headerY = Infinity;

    for (const f of pag.filas) {
      const hasCod = f.celdas.some((c) => /^codigos?$/i.test(normalizarTexto(c.texto)));
      if (!hasCod) continue;
      headerY = f.y;
      for (const c of f.celdas) {
        const n = normalizarTexto(c.texto);
        if (n.includes('precio')) precioX = c.x;
        else if (n === 'c' || n.startsWith('cant')) cantX = c.x;
        else if (n.includes('subtotal')) subtotalX = c.x;
        else if (n.includes('detalle') || n.includes('talles')) notaX = c.x;
      }
      break;
    }

    const filasConCodigo: FilaPedidoRaw[] = [];
    const filasSoloTexto: Array<{ y: number; texto: string }> = [];

    for (const f of pag.filas) {
      if (f.y >= headerY - 1) continue;
      const textoCompleto = f.celdas.map((c) => c.texto).join(' ');
      if (/total\s+sin\s+iva/i.test(textoCompleto)) {
        const numCell = [...f.celdas].reverse().find((c) => /\d/.test(c.texto));
        if (numCell) {
          const val = parsearNumeroDocumento(numCell.texto);
          if (val > 0) totalSinIva = val;
        }
        continue;
      }

      const codCells = f.celdas.filter(
        (c) => c.x < precioX - 60 && /^\d{4}\s*(?:\([^)]*\))?$/.test(c.texto.trim())
      );

      if (codCells.length > 0) {
        const codigos: string[] = [];
        for (const cc of codCells) {
          const m = cc.texto.match(/\b(\d{4})\b/);
          if (m && m[1]) codigos.push(m[1]);
        }

        const qtyCandidates = f.celdas.filter(
          (c) => Math.abs(c.x - cantX) <= 15 && /^\d+$/.test(c.texto.trim())
        );
        qtyCandidates.sort((a, b) => Math.abs(a.x - cantX) - Math.abs(b.x - cantX));
        const qtyCell = qtyCandidates[0];

        const notasCells = f.celdas.filter((c) => c.x >= notaX - 5);
        const numericCells = f.celdas.filter(
          (c) =>
            c !== qtyCell &&
            c.x >= precioX - 30 &&
            c.x < notaX - 5 &&
            c.texto !== '$' &&
            /\d/.test(c.texto)
        );
        const precioCell = numericCells
          .filter((c) => c.x < cantX - 5)
          .sort((a, b) => Math.abs(a.x - precioX) - Math.abs(b.x - precioX))[0];
        const subCell = numericCells
          .filter((c) => c.x >= cantX + 5)
          .sort((a, b) => Math.abs(a.x - subtotalX) - Math.abs(b.x - subtotalX))[0];

        const usados = new Set<CeldaPdf>([
          ...codCells,
          ...notasCells,
          ...(qtyCell ? [qtyCell] : []),
          ...(precioCell ? [precioCell] : []),
          ...(subCell ? [subCell] : [])
        ]);
        const descCells = f.celdas.filter((c) => !usados.has(c) && c.texto !== '$' && c.x < precioX - 20);

        filasConCodigo.push({
          y: f.y,
          codigos,
          descParts: descCells.map((c) => c.texto),
          precio: precioCell ? String(parsearNumeroDocumento(precioCell.texto)) : '',
          qty: qtyCell ? qtyCell.texto.trim() : '',
          subtotal: subCell ? String(parsearNumeroDocumento(subCell.texto)) : '',
          notas: notasCells.map((c) => c.texto.trim())
        });
      } else {
        const descOnly = f.celdas
          .filter((c) => c.x >= 110 && c.x < precioX - 20)
          .map((c) => c.texto)
          .join(' ');
        if (descOnly.trim()) {
          filasSoloTexto.push({ y: f.y, texto: descOnly.trim() });
        }
      }
    }

    for (const extra of filasSoloTexto) {
      let best: FilaPedidoRaw | null = null;
      let bestDist = 14.5;
      for (const fc of filasConCodigo) {
        const d = Math.abs(fc.y - extra.y);
        if (d < bestDist) {
          bestDist = d;
          best = fc;
        }
      }
      if (best) {
        best.descParts.push(extra.texto);
      }
    }

    for (const fc of filasConCodigo) {
      grid.push([
        fc.codigos[0] ?? '',
        fc.codigos[1] ?? '',
        limpiarTextoInvisible(fc.descParts.join(' ')),
        fc.precio ? Number(fc.precio) : null,
        fc.qty ? Number(fc.qty) : null,
        fc.subtotal ? Number(fc.subtotal) : null,
        fc.notas.join(' ')
      ]);
    }
  }

  if (totalSinIva !== null) {
    grid.push(['SUBTOTAL', '', '', null, null, totalSinIva, '']);
  }

  const tab = analizarFilasTabulares(grid, 'PDF', ctx);
  const auditoria = construirAuditoriaDesdeTabular(tab, []);

  return {
    ok: true,
    lineas: tab.lineas,
    advertencias: tab.advertencias,
    avisosPrecio: tab.avisosPrecio,
    tipoArchivo: 'PEDIDO_PDF',
    auditoria,
    correccionesAutomaticas: tab.correcciones,
    hojasDisponibles: [],
    datosDetectados: null,
    comentariosDetectados: tab.comentariosDetectados,
    nombreArchivo: ctx.fileName,
    avisos: [],
    visionDisponible: false,
    requiereConfirmacionHumana: false
  };
};

export const analizarPdf = async (
  bytes: Uint8Array,
  ctx: {
    fileName: string;
    catalogoPorCodigo: ReadonlyMap<string, ProductoCatalogo>;
    catalogoUnificado: ReadonlyArray<ProductoCatalogo>;
  }
): Promise<ResultadoAnalisisArchivo> => {
  let paginas: PaginaPdf[];
  try {
    paginas = await extraerPaginasPdf(bytes);
  } catch {
    return {
      ok: false,
      errorFatal: 'No se pudo leer el archivo PDF.',
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'NO_COMPATIBLE',
      auditoria: auditoriaVacia('BLOQUEADO'),
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados: null,
      comentariosDetectados: [],
      nombreArchivo: ctx.fileName,
      avisos: [],
      visionDisponible: false,
      requiereConfirmacionHumana: false
    };
  }

  const clasif = clasificarPdf(paginas);
  if (clasif === 'IMAGEN_O_PDF_ESCANEADO') {
    const aud = auditoriaVacia('REVISAR');
    aud.motivos.push('PDF escaneado sin texto seleccionable; cargado como referencia visual.');
    return {
      ok: true,
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'IMAGEN_O_PDF_ESCANEADO',
      auditoria: aud,
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados: null,
      comentariosDetectados: [],
      nombreArchivo: ctx.fileName,
      avisos: ['Guardamos este PDF como referencia: podés verlo al lado mientras cargás el pedido.'],
      visionDisponible: false,
      requiereConfirmacionHumana: false
    };
  }

  if (clasif === 'FACTURA_AFIP_PDF') {
    return parsearFacturaAfipDesdePaginas(paginas, ctx);
  }
  if (clasif === 'PEDIDO_PDF') {
    return parsearPedidoPdfDesdePaginas(paginas, ctx);
  }

  return {
    ok: false,
    errorFatal: 'No encontramos artículos ni una factura en este PDF.',
    lineas: [],
    advertencias: [],
    avisosPrecio: [],
    tipoArchivo: 'NO_COMPATIBLE',
    auditoria: auditoriaVacia('BLOQUEADO'),
    correccionesAutomaticas: [],
    hojasDisponibles: [],
    datosDetectados: null,
    comentariosDetectados: [],
    nombreArchivo: ctx.fileName,
    avisos: [],
    visionDisponible: false,
    requiereConfirmacionHumana: false
  };
};
