import * as XLSX from 'xlsx';
import {
  CATALOGO_BY_CODIGO,
  CATALOGO_EXTENDIDO_BY_CODIGO,
  CATALOGO_EXTENDIDO_BY_GRUPO,
  CATALOGO_EXTENDIDO_MIBI
} from './catalogo-mibi';
import { addDaysISO, formatNowDateTime, getTodayISO, recalcularPedido, round2 } from './finanzas';
import { normalizarCuit, validarCuit } from './import/numeros';
import { limpiarTextoComentarioExcel } from './import/xlsx-estilos';
import type {
  AdvertenciaExcel,
  AuditoriaLectura,
  AvisoPrecioExcel,
  CellPrimitive,
  ComentarioCeldaExcel,
  CorreccionAutomatica,
  DatosClienteDetectados,
  DuplicadoEnArchivo,
  EstadoArmado,
  EstadoEnvio,
  HojaDisponible,
  LineaPedido,
  MedioPago,
  Pedido,
  ProductoCatalogo,
  ResultadoParseoLista006,
  SheetJSCellObject,
  SheetJSWorkBook,
  SheetJSWorkSheet,
  SugerenciaProductoCatalogo,
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

const normalizarNotaDesglose = (notaNorm: string): string =>
  notaNorm
    // "1A3 - 2 / 3A6 - 2" o "1 a 3: 2 / 3 a 6: 2" -> "t1 - 2 / t2 - 2"
    .replace(/\b(?:de\s*)?1\s*a\s*3\b/gi, 't1')
    .replace(/\b(?:de\s*)?3\s*a\s*6\b/gi, 't2')
    // "T1 (6562): 4 / T2 (6563): 8" -> "t1 6562: 4 / t2 6563: 8"
    .replace(/\(\s*([a-z0-9\-]+)\s*\)\s*(?::|=|-)?\s*(\d+)/gi, '$1: $2')
    // "6508 - 2 / 6509 - 1" o "t1 - 2 / t2 - 2" -> "6508: 2 / 6509: 1"
    .replace(/\b(\d{4}|t[12])\s*-\s*(\d+)\b/gi, '$1: $2')
    // "6504 X2/6505X2" -> "6504: 2/6505: 2"
    .replace(/\b(\d{4})\s*[x*]\s*(\d+)\b/gi, '$1: $2')
    // "R X2/G X2" o "t1x2" -> "r: 2/g: 2"
    .replace(/\b([grpb]|t[12]|green|verde|pink|rosa|blue|azul|aqua)\s*[x*]\s*(\d+)\b/gi, '$1: $2');

const detectarRepartoEquitativo = (
  notaNorm: string,
  numVariantes: number,
  cantidadTotal: number
): { ok: true; porVariante: number } | { ok: false; motivo: string } | null => {
  if (numVariantes < 2) return null;

  const mEach = notaNorm.match(
    /^(\d+)\s*(?:c\/u\.?|c\.u\.|cada\s*un[oa]|de\s*cada\s*(?:un[oa]|talle|color)|(?:x|por)\s*(?:talle|color|cada|c\/u))$/i
  );
  if (mEach && mEach[1]) {
    const porVar = parseInt(mEach[1], 10);
    if (porVar * numVariantes === cantidadTotal) {
      return { ok: true, porVariante: porVar };
    }
    return { ok: false, motivo: `${porVar} c/u (${porVar * numVariantes}) no coincide con el total ${cantidadTotal}` };
  }

  const esFraseEquitativa =
    /^(?:c\/u\.?|c\.u\.|uno\s*de\s*cada(?:\s*(?:talle|color))?|de\s*cada\s*(?:talle|color|un[oa])|a\s*y\s*r|g\s*y\s*r|t1\s*y\s*t2|amb[oa]s|los\s*dos)$/i.test(
      notaNorm
    ) || /\b(mitad y mitad|50 y 50|50\/50)\b/i.test(notaNorm);

  if (esFraseEquitativa) {
    if (cantidadTotal % numVariantes !== 0) {
      return { ok: false, motivo: `Cantidad (${cantidadTotal}) no divisible equitativamente entre ${numVariantes}` };
    }
    return { ok: true, porVariante: cantidadTotal / numVariantes };
  }

  return null;
};

const extraerCantidadVariante = (
  variante: VarianteCatalogo,
  notaParaParseo: string,
  codigosProducto: Set<string>
): number | null => {
  const aliasList = variante.aliases.length > 0 ? variante.aliases : [variante.codigo, variante.etiqueta];
  const aliasPattern = aliasList.map(escapeRegex).join('|');
  const empiezaConNum = /^\s*\d+/.test(notaParaParseo);

  const reColon = new RegExp(`\\b(?:${aliasPattern})\\b\\s*(?::|=|-)\\s*(\\d+)\\b`, 'gi');
  const reAntes = new RegExp(`\\b(\\d+)\\s*(?:u\\.?|unid\\.?|de)?\\s*\\b(?:${aliasPattern})\\b`, 'gi');
  const reDespues = new RegExp(`\\b(?:${aliasPattern})\\b\\s*(?::|=|-)?\\s*(\\d+)\\b`, 'gi');

  const firstValid = (re: RegExp): number | null => {
    for (const m of notaParaParseo.matchAll(re)) {
      const digits = m[1];
      if (!digits || codigosProducto.has(digits)) continue;
      return parseInt(digits, 10);
    }
    return null;
  };

  const vColon = firstValid(reColon);
  if (vColon !== null) return vColon;

  if (empiezaConNum) {
    return firstValid(reAntes) ?? firstValid(reDespues);
  }
  return firstValid(reDespues) ?? firstValid(reAntes);
};

export const resolverLinea = (
  producto: ProductoCatalogo,
  cantidadTotal: number,
  notaClienteRaw: string,
  idLineaSuffix = ''
): LineaPedido => {
  const notaLimpia = limpiarTextoInvisible(notaClienteRaw);
  const notaNorm = normalizarTexto(notaLimpia);
  const notaParaParseo = normalizarNotaDesglose(notaNorm);
  const idLinea = `${producto.idGrupo}${idLineaSuffix ? `-${idLineaSuffix}` : ''}`;
  const codigosProducto = new Set(producto.variantes.map((v) => v.codigo.toLowerCase()));

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

  // Caso 2: Producto con 3 o más variantes (soporte N variantes)
  if (producto.variantes.length >= 3) {
    const fallbackN = (motivo: string): LineaPedido => ({
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
      desglose: producto.variantes.map((v) => ({ codigo: v.codigo, etiqueta: v.etiqueta, cantidad: 0 }))
    });

    if (!notaNorm) {
      return fallbackN('Sin detalle de talle o color');
    }

    const eqN = detectarRepartoEquitativo(notaNorm, producto.variantes.length, cantidadTotal);
    if (eqN) {
      if (!eqN.ok) return fallbackN(eqN.motivo);
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
        desglose: producto.variantes.map((v) => ({
          codigo: v.codigo,
          etiqueta: v.etiqueta,
          cantidad: eqN.porVariante
        }))
      };
    }

    if (/\b(surtid[oa]s?|variad[oa]s?|eleccion|mezclad[oa]s?|cualquiera)\b/.test(notaNorm)) {
      return fallbackN('Nota ambigua ("surtido" o similar)');
    }

    const cantidadesN = producto.variantes.map((v) => extraerCantidadVariante(v, notaParaParseo, codigosProducto));
    if (cantidadesN.some((c) => c !== null)) {
      const sum = cantidadesN.reduce<number>((acc, c) => acc + (c ?? 0), 0);
      if (sum === cantidadTotal) {
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
          desglose: producto.variantes.map((v, idx) => ({
            codigo: v.codigo,
            etiqueta: v.etiqueta,
            cantidad: cantidadesN[idx] ?? 0
          }))
        };
      }
      return fallbackN(`Las cantidades (${cantidadesN.map((c) => c ?? 0).join(' + ')}) no suman ${cantidadTotal}`);
    }

    const notaSinCodigosN = producto.variantes.reduce(
      (txt, v) => txt.replace(new RegExp(`\\b${escapeRegex(v.codigo)}\\b`, 'gi'), ' '),
      notaNorm
    );
    const mencionaVariantes = producto.variantes.map((v) =>
      v.aliases.some((al) => new RegExp(`\\b${escapeRegex(al)}\\b`, 'i').test(notaNorm))
    );
    const mencionesCount = mencionaVariantes.filter(Boolean).length;
    if (mencionesCount === 1 && !/\d/.test(notaSinCodigosN)) {
      const idxUnico = mencionaVariantes.findIndex(Boolean);
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
        desglose: producto.variantes.map((v, idx) => ({
          codigo: v.codigo,
          etiqueta: v.etiqueta,
          cantidad: idx === idxUnico ? cantidadTotal : 0
        }))
      };
    }

    return fallbackN(`No se pudo determinar el desglose exacto ("${notaLimpia}")`);
  }

  // Caso 3: Producto con 2 variantes (talle o color) -> cero suposiciones
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

  const eq2 = detectarRepartoEquitativo(notaNorm, 2, cantidadTotal);
  if (eq2) {
    if (!eq2.ok) return fallbackRevision(eq2.motivo);
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
        { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: eq2.porVariante },
        { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: eq2.porVariante }
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

  if (/\b(?:de\s*)?1\s*a\s*3\b/.test(notaNorm) && /\b(?:de\s*)?3\s*a\s*6\b/.test(notaNorm)) {
    const qtyT1R = extraerCantidadVariante(v0, notaParaParseo, codigosProducto);
    const qtyT2R = extraerCantidadVariante(v1, notaParaParseo, codigosProducto);
    if (qtyT1R !== null && qtyT2R !== null && qtyT1R + qtyT2R === cantidadTotal) {
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
          { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: qtyT1R },
          { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: qtyT2R }
        ]
      };
    }
  }

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

  // Intentar primero desglose explícito por códigos/aliases para tolerar textos descriptivos ("Pink / rojo")
  let c0 = extraerCantidadVariante(v0, notaParaParseo, codigosProducto);
  let c1 = extraerCantidadVariante(v1, notaParaParseo, codigosProducto);

  // Recuperar typo de 1 dígito en nota con par explícito ("6585: 2/6587: 2" en par 6586-6587)
  if ((c0 === null) !== (c1 === null)) {
    const missingVar = c0 === null ? v0 : v1;
    const matchedVar = c0 === null ? v1 : v0;
    const matchedQty = (c0 ?? c1) ?? 0;
    if (matchedQty > 0 && matchedQty <= cantidadTotal) {
      for (const mTok of notaParaParseo.matchAll(/\b(\d{4})\b\s*(?::|=|-)\s*(\d+)\b/g)) {
        const tok = mTok[1] ?? '';
        const qTok = parseInt(mTok[2] ?? '0', 10);
        if (
          tok &&
          tok !== matchedVar.codigo &&
          distanciaEdicion(tok, missingVar.codigo.toLowerCase()) <= 1 &&
          qTok + matchedQty === cantidadTotal
        ) {
          if (c0 === null) c0 = qTok;
          else c1 = qTok;
          break;
        }
      }
    }
  }

  if (c0 !== null && c1 !== null && c0 + c1 === cantidadTotal && c0 >= 0 && c1 >= 0) {
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
        { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: c0 },
        { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: c1 }
      ]
    };
  }

  const aliasesTodos = new Set([...v0.aliases, ...v1.aliases]);
  for (const colorInv of COLORES_EXTRA_INVALIDOS) {
    const reColor = new RegExp(`\\b${colorInv}\\b`, 'i');
    if (reColor.test(notaNorm) && !aliasesTodos.has(colorInv)) {
      return fallbackRevision(`Color "${colorInv}" no disponible en este modelo`);
    }
  }

  const notaSinRangosNiCodigos = notaNorm
    .replace(/\b(?:de\s*)?1\s*a\s*3(?:\s*anos)?\b/g, ' ')
    .replace(/\b(?:de\s*)?3\s*a\s*6(?:\s*anos)?\b/g, ' ')
    .replace(/\btalle\s*[12]\b/g, ' ')
    .replace(/\bt[12]\b/g, ' ')
    .replace(new RegExp(`\\b${escapeRegex(v0.codigo)}\\b`, 'gi'), ' ')
    .replace(new RegExp(`\\b${escapeRegex(v1.codigo)}\\b`, 'gi'), ' ');
  const numerosSueltos = notaSinRangosNiCodigos.match(/\d+/g);

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
  q0OrCantidades: number | ReadonlyArray<number>,
  q1OrCatalogo?: number | ReadonlyArray<ProductoCatalogo>
): LineaPedido | null => {
  const catalogoExtra = Array.isArray(q1OrCatalogo) ? q1OrCatalogo : [];
  const prod =
    catalogoExtra.find((p) => p.idGrupo === linea.idGrupo) ?? CATALOGO_EXTENDIDO_BY_GRUPO.get(linea.idGrupo);
  const variantesBase: ReadonlyArray<{ codigo: string; etiqueta: string }> = prod
    ? prod.variantes
    : linea.desglose;
  if (variantesBase.length < 2) return null;

  const cantidades: number[] = Array.isArray(q0OrCantidades)
    ? [...q0OrCantidades]
    : [q0OrCantidades, typeof q1OrCatalogo === 'number' ? q1OrCatalogo : 0];

  if (cantidades.length !== variantesBase.length) return null;
  if (cantidades.some((q) => !Number.isInteger(q) || q < 0)) return null;

  const suma = cantidades.reduce((acc, q) => acc + q, 0);
  if (suma !== linea.cantidadTotal) return null;

  return {
    ...linea,
    estado: 'OK',
    requiereRevision: false,
    motivoRevision: undefined,
    desglose: variantesBase.map((v, idx) => ({
      codigo: v.codigo,
      etiqueta: v.etiqueta,
      cantidad: cantidades[idx] ?? 0
    }))
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

export const esFormulaSinResultado = (
  cell: SheetJSCellObject | XLSX.CellObject | undefined
): boolean => {
  if (!cell || typeof cell.f !== 'string' || cell.f.trim().length === 0) return false;
  if (cell.t === 'e') return false;
  if (cell.t === 'z') return true;
  if (cell.v === undefined || cell.v === null) return true;
  if (typeof cell.v === 'string' && cell.v.trim() === '') return true;
  return false;
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
      if (cell && cell.t === 'e') {
        const errTxt = cell.w ? limpiarTextoInvisible(cell.w) : typeof cell.v === 'string' && cell.v ? limpiarTextoInvisible(cell.v) : '#VALUE!';
        row.push(errTxt);
      } else if (esFormulaSinResultado(cell)) {
        row.push(null);
      } else if (cell && cell.v !== undefined) {
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

interface MetadatosHojaExcel {
  filasOcultas: ReadonlySet<number>;
  erroresFormula: ReadonlyMap<string, string>;
  formulasSinResultado: ReadonlyMap<string, string>;
  comentariosCelda: ReadonlyMap<string, string>;
}

const RE_ERROR_FORMULA_TEXTO = /^#(?:VALUE!|REF!|NAME\?|N\/A|DIV\/0!|NUM!)$/i;

const extraerMetadatosHoja = (
  ws: SheetJSWorkSheet,
  comentariosThreaded?: ReadonlyMap<string, string>
): MetadatosHojaExcel => {
  const filasOcultas = new Set<number>();
  const rowsMeta = ws['!rows'];
  if (Array.isArray(rowsMeta)) {
    for (let r = 0; r < rowsMeta.length; r++) {
      if (rowsMeta[r]?.hidden === true) {
        filasOcultas.add(r);
      }
    }
  }

  const erroresFormula = new Map<string, string>();
  const formulasSinResultado = new Map<string, string>();
  const comentariosCelda = new Map<string, string>();

  if (comentariosThreaded) {
    for (const [coord, txt] of comentariosThreaded.entries()) {
      if (txt) comentariosCelda.set(coord, txt);
    }
  }

  for (const [addr, rawCell] of Object.entries(ws)) {
    if (addr.startsWith('!') || !rawCell || typeof rawCell !== 'object' || Array.isArray(rawCell)) {
      continue;
    }
    const cell = rawCell as SheetJSCellObject;
    let decoded: { r: number; c: number } | null = null;
    const ensureDecoded = (): { r: number; c: number } => {
      if (!decoded) decoded = XLSX.utils.decode_cell(addr);
      return decoded;
    };

    if (cell.t === 'e' || (typeof cell.v === 'string' && RE_ERROR_FORMULA_TEXTO.test(cell.v.trim()))) {
      const rc = ensureDecoded();
      const errTxt = cell.w
        ? limpiarTextoInvisible(cell.w)
        : typeof cell.v === 'string' && cell.v
          ? limpiarTextoInvisible(cell.v)
          : '#VALUE!';
      erroresFormula.set(`${rc.r}:${rc.c}`, errTxt);
    } else if (esFormulaSinResultado(cell)) {
      const rc = ensureDecoded();
      formulasSinResultado.set(`${rc.r}:${rc.c}`, `=${cell.f!.trim().replace(/^=+/, '')}`);
    }

    if (Array.isArray(cell.c) && cell.c.length > 0) {
      const partes = cell.c
        .map((cm) => limpiarTextoComentarioExcel(cm.t ?? ''))
        .filter(Boolean);
      if (partes.length > 0) {
        const rc = ensureDecoded();
        const coord = `${rc.r}:${rc.c}`;
        const textoLimpio = partes.join(' · ');
        const existente = comentariosCelda.get(coord);
        if (!existente) {
          comentariosCelda.set(coord, textoLimpio);
        } else if (!existente.includes(textoLimpio)) {
          comentariosCelda.set(coord, `${existente} · ${textoLimpio}`);
        }
      }
    }
  }

  return {
    filasOcultas,
    erroresFormula,
    formulasSinResultado,
    comentariosCelda
  };
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
  return XLSX.read(buffer, { type: 'array', cellStyles: true }) as unknown as SheetJSWorkBook;
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

const STOPWORDS_SUGERENCIA = new Set([
  'para',
  'nino',
  'nina',
  'ninos',
  'ninas',
  'talle',
  'talles',
  'color',
  'colores',
  'unidad',
  'unidades',
  'lista',
  'mibi',
  'nuevo',
  'nueva',
  'nuevos',
  'nuevas',
  'producto',
  'productos',
  'articulo',
  'articulos',
  'detalle',
  'modelo'
]);

const PALABRAS_NO_CODIGO = new Set([
  'TOTAL',
  'SUBTOTAL',
  'ENVIO',
  'DESCUENTO',
  'IVA',
  'NETO',
  'LISTA',
  'CODIGOS',
  'CODIGO',
  'PRODUCTO',
  'CANTIDAD',
  'PRECIO',
  'EFECTIVO',
  'TRANSFERENCIA',
  'CHEQUE'
]);

const distanciaEdicion = (a: string, b: string): number => {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev: number[] = Array.from({ length: n + 1 }, (_, j) => j);
  let curr: number[] = new Array<number>(n + 1).fill(0);
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(
        (prev[j] ?? m + n) + 1,
        (curr[j - 1] ?? m + n) + 1,
        (prev[j - 1] ?? m + n) + cost
      );
    }
    const tmp = prev;
    prev = curr;
    curr = tmp;
  }
  return prev[n] ?? m + n;
};

export const sugerirProductosCercanos = (
  codigoRaw: string,
  descripcionExcel: string,
  catalogo: ReadonlyArray<ProductoCatalogo> = CATALOGO_EXTENDIDO_MIBI
): SugerenciaProductoCatalogo[] => {
  const codTokens = limpiarTextoInvisible(codigoRaw)
    .toUpperCase()
    .split(/[\s\/,|]+/)
    .map((t) => t.trim())
    .filter((t) => t.length >= 3);

  const descKeywords = normalizarTexto(descripcionExcel)
    .split(/[^a-z0-9]+/)
    .filter((w) => w.length >= 4 && !STOPWORDS_SUGERENCIA.has(w));
  const unicosDesc = Array.from(new Set(descKeywords));

  const candidatos: Array<{ score: number; item: SugerenciaProductoCatalogo }> = [];

  for (const prod of catalogo) {
    let bestCodeScore = 0;
    for (const v of prod.variantes) {
      const vCod = v.codigo.trim().toUpperCase();
      if (!vCod || vCod.length < 3) continue;
      for (const c of codTokens) {
        if (c === vCod) {
          bestCodeScore = Math.max(bestCodeScore, 110);
          continue;
        }
        const minLen = Math.min(c.length, vCod.length);
        if (minLen >= 3 && distanciaEdicion(c, vCod) <= 1) {
          bestCodeScore = Math.max(bestCodeScore, 100);
        } else if (minLen >= 3 && (vCod.startsWith(c) || c.startsWith(vCod))) {
          bestCodeScore = Math.max(bestCodeScore, 75);
        }
      }
    }

    let nameHits = 0;
    if (unicosDesc.length > 0) {
      const prodWords = new Set(
        normalizarTexto(prod.nombre)
          .split(/[^a-z0-9]+/)
          .filter((w) => w.length >= 4 && !STOPWORDS_SUGERENCIA.has(w))
      );
      for (const kw of unicosDesc) {
        if (prodWords.has(kw)) nameHits++;
      }
    }

    const nameScore = nameHits >= 2 ? 60 + nameHits * 10 : nameHits === 1 ? 30 : 0;
    const baseScore = Math.max(bestCodeScore, nameScore);
    if (baseScore <= 0) continue;

    const motivo = bestCodeScore >= nameScore && bestCodeScore > 0 ? 'Código parecido' : 'Nombre parecido';
    candidatos.push({
      score: baseScore + (nameHits > 0 && bestCodeScore > 0 ? nameHits * 5 : 0),
      item: {
        idGrupo: prod.idGrupo,
        nombre: prod.nombre,
        codigos: prod.variantes.map((v) => v.codigo),
        rubro: prod.rubro,
        motivo
      }
    });
  }

  candidatos.sort((a, b) => b.score - a.score);
  const vistos = new Set<string>();
  const resultado: SugerenciaProductoCatalogo[] = [];
  for (const c of candidatos) {
    if (vistos.has(c.item.idGrupo)) continue;
    vistos.add(c.item.idGrupo);
    resultado.push(c.item);
    if (resultado.length >= 3) break;
  }
  return resultado;
};

export const aplicarLineaAlBorrador = (
  lineasActuales: ReadonlyArray<LineaPedido>,
  producto: ProductoCatalogo,
  cantidadAAgregar: number,
  notaCliente = ''
): LineaPedido[] => {
  const qty = Math.floor(cantidadAAgregar);
  if (!Number.isFinite(qty) || qty <= 0) return [...lineasActuales];

  const notaLimpia = limpiarTextoInvisible(notaCliente);
  const idx = lineasActuales.findIndex((l) => l.idGrupo === producto.idGrupo);
  if (idx === -1) {
    return [...lineasActuales, resolverLinea(producto, qty, notaLimpia)];
  }

  const existente = lineasActuales[idx]!;
  const totalCombinado = existente.cantidadTotal + qty;
  const notaCombinada = [existente.notaCliente, notaLimpia].filter(Boolean).join(' · ');

  let actualizada: LineaPedido;
  if (producto.variantes.length === 1) {
    actualizada = resolverLinea(producto, totalCombinado, notaCombinada);
  } else {
    const parteNueva = resolverLinea(producto, qty, notaLimpia);
    if (
      existente.estado === 'OK' &&
      !existente.requiereRevision &&
      parteNueva.estado === 'OK' &&
      !parteNueva.requiereRevision
    ) {
      const desgloseCombinado = producto.variantes.map((v) => {
        const cantExist = existente.desglose.find((d) => d.codigo === v.codigo)?.cantidad ?? 0;
        const cantNueva = parteNueva.desglose.find((d) => d.codigo === v.codigo)?.cantidad ?? 0;
        return { codigo: v.codigo, etiqueta: v.etiqueta, cantidad: cantExist + cantNueva };
      });
      const sumaDesglose = desgloseCombinado.reduce((acc, d) => acc + d.cantidad, 0);
      if (sumaDesglose === totalCombinado) {
        actualizada = {
          idLinea: existente.idLinea,
          idGrupo: producto.idGrupo,
          nombreProducto: producto.nombre,
          rubro: producto.rubro,
          precioUnitarioNeto: producto.precioNeto,
          cantidadTotal: totalCombinado,
          notaCliente: notaCombinada,
          estado: 'OK',
          requiereRevision: false,
          desglose: desgloseCombinado
        };
      } else {
        actualizada = resolverLinea(producto, totalCombinado, notaCombinada);
      }
    } else {
      const intentada = resolverLinea(producto, totalCombinado, notaCombinada);
      if (parteNueva.requiereRevision && !intentada.requiereRevision) {
        actualizada = {
          ...intentada,
          estado: 'REQUIERE_REVISION',
          requiereRevision: true,
          motivoRevision: parteNueva.motivoRevision ?? 'Definir talle o color para las unidades agregadas',
          desglose: producto.variantes.map((v) => ({ codigo: v.codigo, etiqueta: v.etiqueta, cantidad: 0 }))
        };
      } else {
        actualizada = intentada;
      }
    }
  }

  return lineasActuales.map((l, i) => (i === idx ? actualizada : l));
};

export const auditoriaVacia = (estadoIntegridad: AuditoriaLectura['estadoIntegridad'] = 'EXACTO'): AuditoriaLectura => ({
  unidadesLeidas: 0,
  unidadesDeclaradasArchivo: null,
  diferenciaUnidades: null,
  unidadesEnAlertas: 0,
  subtotalNetoCalculado: 0,
  totalDeclaradoArchivo: null,
  diferenciaImporte: null,
  filasOmitidasConDatos: 0,
  estadoIntegridad,
  motivos: [],
  descuentoDeclaradoPorc: null,
  descuentoDeclaradoMonto: null,
  ivaDeclaradoArchivo: null,
  totalFacturadoDeclaradoArchivo: null,
  duplicadosEnArchivo: []
});

export const detectarNegacionODuda = (texto: string): boolean => {
  if (!texto) return false;
  if (texto.includes('?')) return true;
  const norm = normalizarTexto(texto);
  if (!norm) return false;
  return /\b(?:no|no\s+mandar|no\s+enviar|no\s+incluir|no\s+va|cancelar|cancelado|cancelada|anular|anulado|anulada|sacar|eliminar|sin\s+stock|faltante|consultar|a\s+confirmar|confirmar)\b/i.test(
    norm
  );
};

export const detectarCantidadContradictoriaTalleUnico = (
  texto: string,
  cantidadTotal: number,
  codigosProducto: ReadonlySet<string> = new Set()
): number | null => {
  if (!texto) return null;
  let norm = normalizarTexto(texto);
  if (!norm) return null;

  for (const cod of codigosProducto) {
    const cleanCod = normalizarTexto(cod);
    if (cleanCod) {
      norm = norm.replace(new RegExp(`\\b${escapeRegex(cleanCod)}\\b`, 'gi'), ' ');
    }
  }

  norm = norm
    // Fechas dd/mm o dd/mm/yyyy
    .replace(/\b\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?\b/g, ' ')
    // Medidas 60x60 o 1,50 x 0,90
    .replace(/\b\d+(?:[.,]\d+)?\s*x\s*\d+(?:[.,]\d+)?\b/gi, ' ')
    // Edades 1 a 3 / 3 a 6 / 12-24 meses
    .replace(/\b(?:de\s*)?\d{1,2}\s*(?:a|-)\s*\d{1,2}(?:\s*(?:anos?|meses|m))?\b/gi, ' ')
    // Códigos numéricos de 4 dígitos (6526, 2026, etc.)
    .replace(/\b\d{4}\b/g, ' ')
    // Frases de reparto equitativo ya válidas ("2 de cada talle", "4 c/u")
    .replace(/\b\d+\s*(?:c\/u\.?|c\.u\.|cada\s*un[oa]|de\s*cada\s*(?:un[oa]|talle|color))\b/gi, ' ');

  const patrones: ReadonlyArray<RegExp> = [
    /\b(?:mandar|enviar|solo|solamente|unicamente|en\s+realidad|son|pedir|total)\s+(?:son\s+)?(\d+)\b/gi,
    /\b(\d+)\s*(?:u\.?(?=\s|$|[.,;])|unid(?:ades|ad|\.)?)\b/gi,
    /(?:^|\s)[x*]\s*(\d+)\b/gi
  ];

  for (const re of patrones) {
    for (const m of norm.matchAll(re)) {
      const n = parseInt(m[1] ?? '0', 10);
      if (Number.isInteger(n) && n > 0 && n !== cantidadTotal) {
        return n;
      }
    }
  }

  return null;
};

const STOPWORDS_IDENTIFICACION = new Set([
  'de',
  'del',
  'la',
  'el',
  'los',
  'las',
  'con',
  'sin',
  'para',
  'por',
  'en',
  'y',
  'a',
  't',
  'talle',
  'talles',
  't1',
  't2',
  'unico',
  'kids',
  'neoprene',
  'placa',
  'completa',
  'flotacion',
  'retirable',
  'plastico',
  'plasticos',
  'silicona',
  'lentes',
  'filtro',
  'uv',
  'agua',
  'arena',
  'anos',
  'ano'
]);

const tokenizarSignificativos = (texto: string): string[] =>
  Array.from(
    new Set(
      normalizarTexto(texto)
        .split(/[^a-z0-9]+/)
        .filter((w) => w.length >= 2 && !STOPWORDS_IDENTIFICACION.has(w) && !/^\d+$/.test(w))
    )
  );

export const identificarProductoPorDescripcion = (
  descripcion: string,
  catalogo: ReadonlyArray<ProductoCatalogo> = CATALOGO_EXTENDIDO_MIBI
): ProductoCatalogo | null => {
  const descTokens = new Set(tokenizarSignificativos(descripcion));
  if (descTokens.size === 0) return null;

  const matches: ProductoCatalogo[] = [];
  for (const prod of catalogo) {
    const prodTokens = tokenizarSignificativos(prod.nombre);
    if (prodTokens.length === 0) continue;
    if (prodTokens.every((t) => descTokens.has(t))) {
      matches.push(prod);
    }
  }
  return matches.length === 1 ? matches[0]! : null;
};

export const corregirCodigoConContexto = (
  codigo: string,
  descNorm: string,
  modeloNorm: string
): { corregido: string; motivo: string } | null => {
  if (
    codigo === '6570' &&
    descNorm.includes('chaleco') &&
    (descNorm.includes('bug') || modeloNorm.includes('bug'))
  ) {
    return {
      corregido: '6540',
      motivo: 'Código 6570 con descripción Chaleco Bug corregido automáticamente a 6540'
    };
  }
  return null;
};

export interface ResultadoAnalisisTabular {
  lineas: LineaPedido[];
  advertencias: AdvertenciaExcel[];
  avisosPrecio: AvisoPrecioExcel[];
  correcciones: CorreccionAutomatica[];
  comentariosDetectados: ComentarioCeldaExcel[];
  duplicadosEnArchivo: DuplicadoEnArchivo[];
  erroresFormulaPie: string[];
  formulasSinResultadoPie: string[];
  filasSubtotalFormulaSinResultado: number;
  unidadesDeclaradas: number | null;
  subtotalDeclarado: number | null;
  descuentoDeclaradoPorc: number | null;
  descuentoDeclaradoMonto: number | null;
  ivaDeclaradoArchivo: number | null;
  totalConIvaDeclarado: number | null;
  subtotalCalculado: number;
  filasOmitidasConDatos: number;
  descuentoDetectado: number | null;
  headerRowIdx: number;
}

const RE_FILA_RESUMEN =
  /^(?:sub\s*total|total(?:es)?\b|iva\b|impuestos?|descuentos?|bonificaci|neto\b|importe\s+total|total\s+a\s+pagar|saldo|condiciones\s+de\s+venta)/;

const extractNumericFromCell = (raw: CellPrimitive): number | null => {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'number') return Number.isFinite(raw) ? round2(raw) : null;
  const s = limpiarTextoInvisible(String(raw));
  if (!s) return null;
  const parsed = parsearMontoArgentino(s);
  return Number.isFinite(parsed) && parsed !== 0 ? parsed : /^0+(?:[.,]0+)?$/.test(s.replace(/\$/g, '').trim()) ? 0 : null;
};

const rightmostNumeric = (row: CellPrimitive[]): number | null => {
  for (let c = row.length - 1; c >= 0; c--) {
    const val = extractNumericFromCell(row[c] ?? null);
    if (val !== null) return val;
  }
  return null;
};

export const analizarFilasTabulares = (
  rows: CellPrimitive[][],
  hoja: string,
  ctx: {
    catalogoPorCodigo: ReadonlyMap<string, ProductoCatalogo>;
    catalogoUnificado: ReadonlyArray<ProductoCatalogo>;
    fileName: string;
    filasOcultas?: ReadonlySet<number>;
    celdasTachadas?: ReadonlySet<string>;
    erroresFormula?: ReadonlyMap<string, string>;
    formulasSinResultado?: ReadonlyMap<string, string>;
    comentariosCelda?: ReadonlyMap<string, string>;
  }
): ResultadoAnalisisTabular => {
  const {
    catalogoPorCodigo,
    catalogoUnificado,
    filasOcultas,
    celdasTachadas,
    erroresFormula,
    formulasSinResultado,
    comentariosCelda
  } = ctx;
  const advertencias: AdvertenciaExcel[] = [];
  const avisosPrecioMap = new Map<string, AvisoPrecioExcel>();
  const correcciones: CorreccionAutomatica[] = [];
  const correccionesPendientesLineaId: Array<{ corrIdx: number; idGrupo: string }> = [];
  const comentariosDetectados: ComentarioCeldaExcel[] = [];
  const erroresFormulaPie: string[] = [];
  const formulasSinResultadoPie: string[] = [];
  let filasSubtotalFormulaSinResultado = 0;
  const aparicionesPorCodigo = new Map<
    string,
    { codigo: string; nombreProducto: string; idGrupo: string; apariciones: Array<{ fila: number; cantidad: number }> }
  >();
  let lineasAcumuladas: LineaPedido[] = [];

  let colCodeA = 0;
  let colCodeEnd = 1;
  let colDesc = 2;
  let colPrecio = 3;
  let colQty = 4;
  let colSubtotal = -1;
  let colTotalPed = -1;
  let colT1 = -1;
  let colT2 = -1;
  let colNota = -1;
  let colModelo = -1;
  let headerDetectado = false;
  let hadNotaHeader = false;
  let hadSubtotalHeader = false;
  let hadPrecioHeader = false;
  let headerRowIdx = -1;

  let unidadesDeclaradas: number | null = null;
  let subtotalDeclarado: number | null = null;
  let descuentoDeclaradoPorc: number | null = null;
  let descuentoDeclaradoMonto: number | null = null;
  let ivaDeclaradoArchivo: number | null = null;
  let totalConIvaDeclarado: number | null = null;
  let subtotalCalculado = 0;
  let filasOmitidasConDatos = 0;
  let descuentoDetectado: number | null = null;

  const getCellError = (rIdx: number, cIdx: number): string | null => {
    if (cIdx < 0) return null;
    const metaErr = erroresFormula?.get(`${rIdx}:${cIdx}`);
    if (metaErr) return metaErr;
    const raw = rows[rIdx]?.[cIdx];
    if (typeof raw === 'string' && RE_ERROR_FORMULA_TEXTO.test(limpiarTextoInvisible(raw))) {
      return limpiarTextoInvisible(raw).toUpperCase();
    }
    return null;
  };

  const getCellFormulaSinResultado = (rIdx: number, cIdx: number): string | null => {
    if (cIdx < 0 || !formulasSinResultado) return null;
    return formulasSinResultado.get(`${rIdx}:${cIdx}`) ?? null;
  };

  for (let rIdx = 0; rIdx < rows.length; rIdx++) {
    const row = rows[rIdx];
    if (!row) continue;
    const rowTexts = row.map((c) => (c === null ? '' : normalizarTexto(String(c)).replace(/\s+/g, ' ').trim()));

    const idxCantPed = rowTexts.findIndex(
      (t) => t.includes('cantidades pedidas') || t.includes('cantidad pedida')
    );
    const idxT1 = rowTexts.findIndex((t) => t.includes('pedidas t1') || t === 't1' || t === 'talle 1');
    const idxT2 = rowTexts.findIndex((t) => t.includes('pedidas t2') || t === 't2' || t === 'talle 2');
    const idxTotalPed = rowTexts.findIndex((t) => t.includes('total pedido'));
    const idxCantFallback = rowTexts.findIndex((t) => t === 'cantidad' || t === 'cant');
    const idxSubtotal = rowTexts.findIndex((t) => t === 'subtotal' || t === 'sub total');
    const idxTalleColor = rowTexts.findIndex(
      (t) => t.includes('talles/colores') || t.includes('talles / colores') || t.includes('talles/ colores')
    );
    const idxModelo = rowTexts.findIndex((t) => t.includes('modelo') || t === 'talle');
    const idxObsNota = rowTexts.findIndex(
      (t) =>
        t === 'obs' ||
        t.startsWith('observ') ||
        t === 'nota' ||
        t === 'notas' ||
        t.startsWith('coment') ||
        t.startsWith('aclarac')
    );
    const idxDetalleFallback = rowTexts.findIndex((t) => t.includes('detalle') && !t.includes('producto'));
    const idxCodigos = rowTexts.findIndex((t) => t === 'codigos' || t === 'codigo' || t === 'codigo t1');
    const idxCodigoT2 = rowTexts.findIndex((t) => t === 'codigo t2');
    const idxPrecio = rowTexts.findIndex((t) => t === 'precio' || t.includes('precio neto') || t.includes('precio unit'));
    const idxDesc = rowTexts.findIndex(
      (t) =>
        t === 'producto' ||
        t === 'productos' ||
        t === 'descripcion' ||
        t === 'caracteristicas' ||
        t === 'articulo' ||
        t === 'articulos' ||
        t.includes('detalle producto') ||
        t.includes('detalle del producto')
    );

    if (
      idxCantPed !== -1 ||
      (idxT1 !== -1 && idxT2 !== -1) ||
      (idxCodigos !== -1 && (idxCantFallback !== -1 || idxPrecio !== -1 || idxDesc !== -1))
    ) {
      headerDetectado = true;
      if (headerRowIdx === -1) headerRowIdx = rIdx;
      if (idxCodigos !== -1) {
        colCodeA = idxCodigos;
        if (idxCodigoT2 !== -1) {
          colCodeEnd = idxCodigoT2;
        } else if (idxDesc > idxCodigos + 1) {
          colCodeEnd = idxDesc - 1;
        } else if (idxDesc === idxCodigos + 1) {
          colCodeEnd = idxCodigos;
        } else {
          colCodeEnd = idxCodigos + 1;
        }
      }
      const effDescIdx = idxDesc !== -1 ? idxDesc : idxDetalleFallback;
      if (effDescIdx !== -1) {
        colDesc = effDescIdx;
        if (colCodeEnd >= effDescIdx && effDescIdx > colCodeA) {
          colCodeEnd = effDescIdx - 1;
        }
      }
      if (idxPrecio !== -1) {
        colPrecio = idxPrecio;
        hadPrecioHeader = true;
      } else if (!hadPrecioHeader) {
        colPrecio = -1;
      }
      if (idxSubtotal !== -1) {
        colSubtotal = idxSubtotal;
        hadSubtotalHeader = true;
      } else if (!hadSubtotalHeader) {
        colSubtotal = -1;
      }
      if (idxT1 !== -1 && idxT2 !== -1) {
        colT1 = idxT1;
        colT2 = idxT2;
        colTotalPed = idxTotalPed;
        if (idxTotalPed !== -1) colQty = idxTotalPed;
      } else if (idxCantPed !== -1 || idxCantFallback !== -1 || idxTotalPed !== -1) {
        colT1 = -1;
        colT2 = -1;
        colTotalPed = idxTotalPed;
        if (idxTotalPed !== -1) colQty = idxTotalPed;
        else if (idxCantPed !== -1) colQty = idxCantPed;
        else if (idxCantFallback !== -1) colQty = idxCantFallback;
      }

      if (idxTalleColor !== -1) {
        colNota = idxTalleColor;
        hadNotaHeader = true;
      } else if (idxObsNota !== -1) {
        colNota = idxObsNota;
        hadNotaHeader = true;
      } else if (idxDetalleFallback !== -1 && idxDetalleFallback !== effDescIdx) {
        colNota = idxDetalleFallback;
        hadNotaHeader = true;
      } else if (!hadNotaHeader) {
        colNota = -1;
      }

      if (idxModelo !== -1 && idxModelo !== colNota) colModelo = idxModelo;
      continue;
    }

    const extractCode = (txt: string): string | null => {
      if (!txt || PALABRAS_NO_CODIGO.has(txt)) return null;
      if (
        /\b(?:ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)\b/.test(txt) ||
        /\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(txt) ||
        /^LISTA\b/.test(txt)
      ) {
        return null;
      }
      if (catalogoPorCodigo.has(txt)) return txt;
      const mNum = txt.match(/\b(\d{4})\b/);
      if (mNum && mNum[1]) return mNum[1];
      const mExt = txt.match(/\b([A-Z0-9]{2,12}-[A-Z0-9]{1,12})\b/);
      if (mExt && mExt[1]) return mExt[1];
      for (const token of txt.split(/[\s\/,|]+/)) {
        if (token && catalogoPorCodigo.has(token)) return token;
      }
      if (!txt.includes(' ') && /^[A-Z0-9]{3,16}$/.test(txt) && /(?:\d.*){2,}/.test(txt)) {
        return txt;
      }
      return null;
    };

    // Verificar si la fila es de resumen antes de extraer códigos
    const resumenTexto = rowTexts.find((t) => RE_FILA_RESUMEN.test(t));
    let tieneCodigoCatalogoEnRango = false;
    if (resumenTexto) {
      for (let cIdx = Math.max(0, colCodeA); cIdx <= Math.min(row.length - 1, colCodeEnd); cIdx++) {
        const rawCell = row[cIdx];
        if (rawCell === null || rawCell === undefined) continue;
        const cleanCell = limpiarTextoInvisible(String(rawCell)).toUpperCase();
        if (!cleanCell) continue;
        for (const c4 of cleanCell.matchAll(/\b(\d{4})\b/g)) {
          if (c4[1] && catalogoPorCodigo.has(c4[1])) {
            tieneCodigoCatalogoEnRango = true;
            break;
          }
        }
        if (tieneCodigoCatalogoEnRango) break;
        const ext = extractCode(cleanCell);
        if (ext && catalogoPorCodigo.has(ext.toUpperCase())) {
          tieneCodigoCatalogoEnRango = true;
          break;
        }
      }
    }

    if (resumenTexto && !tieneCodigoCatalogoEnRango) {
      for (let cErr = 0; cErr < row.length; cErr++) {
        const errPie = getCellError(rIdx, cErr);
        if (errPie) {
          erroresFormulaPie.push(errPie);
          break;
        }
        const formulaSinResPie = getCellFormulaSinResultado(rIdx, cErr);
        if (formulaSinResPie) {
          formulasSinResultadoPie.push(formulaSinResPie);
        }
      }

      if (
        unidadesDeclaradas === null &&
        (/^(?:total|total\s+(?:de\s+)?(?:unidades|cantidad(?:es)?|u\.?)|total\s+pedido)$/.test(resumenTexto) ||
          /^total\s+de\s+unidades\b/.test(resumenTexto))
      ) {
        if (colTotalPed !== -1) {
          const qp = parsearCantidadEstricta(row[colTotalPed] ?? null);
          if (qp.valid && qp.qty > 0) unidadesDeclaradas = qp.qty;
        } else if (colT1 !== -1 && colT2 !== -1) {
          const p1 = parsearCantidadEstricta(row[colT1] ?? null);
          const p2 = parsearCantidadEstricta(row[colT2] ?? null);
          if (p1.valid && p2.valid && p1.qty + p2.qty > 0) {
            unidadesDeclaradas = p1.qty + p2.qty;
          }
        } else {
          const qc = parsearCantidadEstricta(headerDetectado ? (row[colQty] ?? null) : (row[4] ?? null));
          if (qc.valid && qc.qty > 0) {
            unidadesDeclaradas = qc.qty;
          }
        }
      } else if (/^(?:sub\s*total|neto|total\s+b)\b/.test(resumenTexto)) {
        if (subtotalDeclarado === null) {
          const subCell =
            colSubtotal >= 0
              ? extractNumericFromCell(row[colSubtotal] ?? null)
              : headerDetectado
                ? null
                : extractNumericFromCell(row[5] ?? null);
          const subVal = subCell ?? rightmostNumeric(row);
          if (subVal !== null && subVal > 0) {
            subtotalDeclarado = round2(subVal);
          }
        }
      } else if (/^(?:total\s+facturado|total\s+final|importe\s+total|total\s+a\s+pagar)\b/.test(resumenTexto)) {
        if (totalConIvaDeclarado === null) {
          const totCell =
            colSubtotal >= 0
              ? extractNumericFromCell(row[colSubtotal] ?? null)
              : headerDetectado
                ? null
                : extractNumericFromCell(row[5] ?? null);
          const totVal = totCell ?? rightmostNumeric(row);
          if (totVal !== null && totVal > 0) {
            totalConIvaDeclarado = round2(totVal);
          }
        }
      } else if (/^(?:iva\b|impuestos?)/.test(resumenTexto)) {
        if (ivaDeclaradoArchivo === null) {
          const ivaCell =
            colSubtotal >= 0
              ? extractNumericFromCell(row[colSubtotal] ?? null)
              : headerDetectado
                ? null
                : extractNumericFromCell(row[5] ?? null);
          const ivaRight = rightmostNumeric(row);
          const ivaVal = ivaCell ?? (ivaRight !== 21 ? ivaRight : null);
          if (ivaVal !== null && ivaVal >= 0) {
            ivaDeclaradoArchivo = round2(ivaVal);
          }
        }
      } else if (/^(?:descuento|bonificaci)/.test(resumenTexto)) {
        for (const t of rowTexts) {
          const mPct = t.match(/(\d+(?:[.,]\d+)?)\s*%/);
          if (mPct && mPct[1]) {
            const pct = parseFloat(mPct[1].replace(',', '.'));
            if (Number.isFinite(pct) && pct > 0 && pct <= 50) {
              if (descuentoDeclaradoPorc === null) descuentoDeclaradoPorc = round2(pct);
              if (descuentoDetectado === null) descuentoDetectado = round2(pct);
              break;
            }
          }
        }
        if (descuentoDeclaradoMonto === null) {
          const descSubCell =
            colSubtotal >= 0
              ? extractNumericFromCell(row[colSubtotal] ?? null)
              : headerDetectado
                ? null
                : extractNumericFromCell(row[5] ?? null);
          const descVal = descSubCell !== null ? Math.abs(descSubCell) : null;
          if (descVal !== null && descVal > 0) {
            descuentoDeclaradoMonto = round2(descVal);
          }
        }
      }
      continue;
    }

    const codigosEnFila: string[] = [];
    for (let cIdx = Math.max(0, colCodeA); cIdx <= Math.min(row.length - 1, colCodeEnd); cIdx++) {
      const rawCell = row[cIdx];
      if (rawCell === null || rawCell === undefined) continue;
      const cleanCell = limpiarTextoInvisible(String(rawCell)).toUpperCase();
      if (!cleanCell) continue;
      if (
        /\b(?:ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)\b/.test(cleanCell) ||
        /\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(cleanCell) ||
        /^LISTA\b/.test(cleanCell)
      ) {
        continue;
      }
      for (const c4 of cleanCell.matchAll(/\b(\d{4})\b/g)) {
        if (c4[1] && !codigosEnFila.includes(c4[1])) codigosEnFila.push(c4[1]);
      }
      const extracted = extractCode(cleanCell);
      if (extracted && !codigosEnFila.includes(extracted)) {
        codigosEnFila.push(extracted);
      }
    }

    const descCell = headerDetectado ? (row[colDesc] ?? null) : (row[2] ?? null);
    const precioCell = headerDetectado ? (colPrecio >= 0 ? (row[colPrecio] ?? null) : null) : (row[3] ?? null);
    const notaCell = headerDetectado ? (colNota >= 0 ? (row[colNota] ?? null) : null) : (row[6] ?? null);
    const modeloCell = headerDetectado && colModelo >= 0 ? (row[colModelo] ?? null) : null;
    const subtotalCell = headerDetectado
      ? colSubtotal >= 0
        ? (row[colSubtotal] ?? null)
        : null
      : (row[5] ?? null);

    const usaT1T2SinTotal = colT1 !== -1 && colT2 !== -1 && colTotalPed === -1;
    const rawT1 = colT1 !== -1 ? (row[colT1] ?? null) : null;
    const rawT2 = colT2 !== -1 ? (row[colT2] ?? null) : null;
    const qtyCell = usaT1T2SinTotal
      ? (() => {
          const p1 = parsearCantidadEstricta(rawT1);
          const p2 = parsearCantidadEstricta(rawT2);
          if (p1.valid && p2.valid) return p1.qty + p2.qty;
          return rawT1 ?? rawT2;
        })()
      : headerDetectado
        ? (row[colQty] ?? null)
        : (row[4] ?? null);

    const descripcionExcel =
      descCell !== null && descCell !== undefined ? limpiarTextoInvisible(String(descCell)) : '';
    const notaDetectadaRaw =
      notaCell !== null && notaCell !== undefined ? limpiarTextoInvisible(String(notaCell)) : '';
    const modeloDetectado =
      modeloCell !== null && modeloCell !== undefined ? limpiarTextoInvisible(String(modeloCell)) : '';
    const cantidadRaw =
      qtyCell !== null && qtyCell !== undefined ? limpiarTextoInvisible(String(qtyCell)) : '';

    // Comentarios flotantes o modernos en cualquier celda relevante de esta fila
    const colsRelevantesSet = new Set<number>();
    for (let cIdx = Math.max(0, colCodeA); cIdx <= Math.min(row.length - 1, colCodeEnd); cIdx++) {
      colsRelevantesSet.add(cIdx);
    }
    colsRelevantesSet.add(headerDetectado ? colDesc : 2);
    if (headerDetectado ? colQty >= 0 : true) colsRelevantesSet.add(headerDetectado ? colQty : 4);
    if (colT1 >= 0) colsRelevantesSet.add(colT1);
    if (colT2 >= 0) colsRelevantesSet.add(colT2);
    if (headerDetectado && colNota >= 0) colsRelevantesSet.add(colNota);
    else if (!headerDetectado) colsRelevantesSet.add(6);

    const rowComentarios: Array<{ col: number; texto: string }> = [];
    if (comentariosCelda && comentariosCelda.size > 0) {
      for (const cIdx of Array.from(colsRelevantesSet).sort((a, b) => a - b)) {
        const cmTxt = comentariosCelda.get(`${rIdx}:${cIdx}`);
        if (cmTxt) {
          rowComentarios.push({ col: cIdx, texto: cmTxt });
        }
      }
    }
    const comentarioFilaTexto = Array.from(new Set(rowComentarios.map((x) => x.texto))).join(' · ');

    // Verificar celdas tachadas en columnas clave de artículo
    let filaEstaTachada = false;
    if (celdasTachadas && celdasTachadas.size > 0) {
      const colsStrikeCheck: number[] = [];
      for (let cIdx = Math.max(0, colCodeA); cIdx <= Math.min(row.length - 1, colCodeEnd); cIdx++) {
        colsStrikeCheck.push(cIdx);
      }
      colsStrikeCheck.push(headerDetectado ? colDesc : 2);
      colsStrikeCheck.push(headerDetectado ? colQty : 4);
      if (colT1 >= 0) colsStrikeCheck.push(colT1);
      if (colT2 >= 0) colsStrikeCheck.push(colT2);
      for (const cIdx of colsStrikeCheck) {
        if (cIdx >= 0 && celdasTachadas.has(`${rIdx}:${cIdx}`)) {
          filaEstaTachada = true;
          break;
        }
      }
    }

    // Blindaje 5A: Errores de fórmula (#VALUE!, #REF!, etc.) en celdas numéricas de artículo
    const colsErrorCheck: number[] = usaT1T2SinTotal
      ? [colT1, colT2, headerDetectado ? colPrecio : 3, headerDetectado ? colSubtotal : 5]
      : [
          headerDetectado ? colQty : 4,
          headerDetectado ? colPrecio : 3,
          headerDetectado ? colSubtotal : 5
        ];
    let errorFormulaFila: string | null = null;
    for (const cIdx of colsErrorCheck) {
      const err = getCellError(rIdx, cIdx);
      if (err) {
        errorFormulaFila = err;
        break;
      }
    }

    if (errorFormulaFila && (codigosEnFila.length > 0 || (headerDetectado && rowTexts.some((t) => (t.match(/[a-z]/g)?.length ?? 0) >= 3)))) {
      const codJoinedErr = codigosEnFila.join('/');
      let prodErr: ProductoCatalogo | undefined;
      for (const cod of codigosEnFila) {
        const f =
          catalogoPorCodigo.get(cod.toUpperCase()) ??
          CATALOGO_EXTENDIDO_BY_CODIGO.get(cod) ??
          CATALOGO_BY_CODIGO.get(cod);
        if (f) {
          prodErr = f;
          break;
        }
      }
      filasOmitidasConDatos += 1;
      advertencias.push({
        hoja,
        fila: rIdx + 1,
        codigo: codJoinedErr,
        mensaje: `Error de fórmula (${errorFormulaFila}) en la fila.`,
        id: `${hoja}#${rIdx + 1}#${codJoinedErr || 'SIN_COD'}#ERROR_FORMULA`,
        tipo: 'ERROR_FORMULA',
        codigosDetectados: codigosEnFila,
        descripcionExcel,
        precioExcelReferencia: null,
        cantidadDetectada: null,
        cantidadRaw: errorFormulaFila,
        notaDetectada: notaDetectadaRaw || modeloDetectado,
        productoExistenteIdGrupo: prodErr ? prodErr.idGrupo : null,
        sugerenciasCatalogo: prodErr
          ? []
          : sugerirProductosCercanos(codJoinedErr, descripcionExcel, catalogoUnificado)
      });
      continue;
    }

    let precioExcelReferencia: number | null = null;
    if (precioCell !== null && precioCell !== undefined && limpiarTextoInvisible(String(precioCell)) !== '') {
      const parsedPrecio =
        typeof precioCell === 'number'
          ? Number.isFinite(precioCell)
            ? round2(precioCell)
            : 0
          : parsearMontoArgentino(precioCell);
      if (parsedPrecio > 0) {
        precioExcelReferencia = parsedPrecio;
      }
    }

    // Micro-Blindaje A: Fórmula de Excel sin valor precalculado en celdas de cantidad
    const colsQtyFormulaCheck: number[] =
      colT1 !== -1 && colT2 !== -1
        ? colTotalPed !== -1
          ? [colTotalPed, colT1, colT2]
          : [colT1, colT2]
        : [headerDetectado ? colQty : 4];
    let formulaSinResultadoQty: string | null = null;
    for (const cIdx of colsQtyFormulaCheck) {
      const fSin = getCellFormulaSinResultado(rIdx, cIdx);
      if (fSin) {
        formulaSinResultadoQty = fSin;
        break;
      }
    }
    if (
      formulaSinResultadoQty &&
      (codigosEnFila.length > 0 ||
        (headerDetectado &&
          (rowTexts.some((t) => (t.match(/[a-z]/g)?.length ?? 0) >= 3) ||
            precioExcelReferencia !== null)))
    ) {
      const descNormFormula = normalizarTexto(descripcionExcel);
      const modeloNormFormula = normalizarTexto(modeloDetectado);
      const codigosCorregidosFormula = codigosEnFila.map(
        (c) => corregirCodigoConContexto(c, descNormFormula, modeloNormFormula)?.corregido ?? c
      );
      const codJoinedFormula = codigosCorregidosFormula.join('/');
      let prodFormula: ProductoCatalogo | undefined;
      for (const cod of codigosCorregidosFormula) {
        const enc =
          catalogoPorCodigo.get(cod.toUpperCase()) ??
          CATALOGO_EXTENDIDO_BY_CODIGO.get(cod) ??
          CATALOGO_BY_CODIGO.get(cod);
        if (enc) {
          prodFormula = enc;
          break;
        }
      }
      if (!prodFormula && descripcionExcel) {
        prodFormula = identificarProductoPorDescripcion(descripcionExcel, catalogoUnificado) ?? undefined;
      }
      filasOmitidasConDatos += 1;
      advertencias.push({
        hoja,
        fila: rIdx + 1,
        codigo: codJoinedFormula,
        mensaje:
          'Esta celda tiene una fórmula sin resultado guardado. Ingresá la cantidad real o guardá nuevamente el Excel.',
        id: `${hoja}#${rIdx + 1}#${codJoinedFormula || 'SIN_COD'}#ERROR_FORMULA_SIN_RESULTADO`,
        tipo: 'ERROR_FORMULA',
        codigosDetectados: codigosCorregidosFormula,
        descripcionExcel,
        precioExcelReferencia,
        cantidadDetectada: null,
        cantidadRaw: formulaSinResultadoQty,
        notaDetectada: notaDetectadaRaw || comentarioFilaTexto || modeloDetectado,
        productoExistenteIdGrupo: prodFormula ? prodFormula.idGrupo : null,
        sugerenciasCatalogo: prodFormula
          ? []
          : sugerirProductosCercanos(codJoinedFormula, descripcionExcel, catalogoUnificado)
      });
      continue;
    }

    if (codigosEnFila.length === 0) {
      if (!headerDetectado) continue;
      const qtyCheck = parsearCantidadEstricta(qtyCell);
      const tieneCantidadOTextoDigitos =
        (qtyCheck.valid && qtyCheck.qty > 0) || (!qtyCheck.valid && /\d/.test(cantidadRaw));
      const tieneTextoConLetras = rowTexts.some((t) => (t.match(/[a-z]/g)?.length ?? 0) >= 3);
      const tienePrecioNum = precioExcelReferencia !== null && precioExcelReferencia > 0;
      if (!tieneCantidadOTextoDigitos || (!tieneTextoConLetras && !tienePrecioNum)) {
        continue;
      }

      const qtyParaAlerta = qtyCheck.valid ? qtyCheck.qty : null;
      const matchUnico = identificarProductoPorDescripcion(descripcionExcel, catalogoUnificado);
      const cercanas = sugerirProductosCercanos('', descripcionExcel, catalogoUnificado);
      const sugerencias: SugerenciaProductoCatalogo[] = [];
      if (matchUnico) {
        sugerencias.push({
          idGrupo: matchUnico.idGrupo,
          nombre: matchUnico.nombre,
          codigos: matchUnico.variantes.map((v) => v.codigo),
          rubro: matchUnico.rubro,
          motivo: 'Coincidencia por descripción'
        });
      }
      for (const s of cercanas) {
        if (!sugerencias.some((x) => x.idGrupo === s.idGrupo)) {
          sugerencias.push(s);
        }
      }

      if (filasOcultas?.has(rIdx) && qtyParaAlerta !== null && qtyParaAlerta > 0) {
        filasOmitidasConDatos += 1;
        advertencias.push({
          hoja,
          fila: rIdx + 1,
          codigo: '',
          mensaje: `Esta fila está oculta en el Excel pero tiene ${qtyParaAlerta} u. cargadas.`,
          id: `${hoja}#${rIdx + 1}#FILA_OCULTA`,
          tipo: 'FILA_OCULTA',
          codigosDetectados: [],
          descripcionExcel,
          precioExcelReferencia,
          cantidadDetectada: qtyParaAlerta,
          cantidadRaw,
          notaDetectada: notaDetectadaRaw || modeloDetectado,
          productoExistenteIdGrupo: matchUnico ? matchUnico.idGrupo : null,
          sugerenciasCatalogo: sugerencias.slice(0, 4)
        });
        continue;
      }

      const subFilaVal = extractNumericFromCell(subtotalCell);
      if (subFilaVal !== null && subFilaVal > 0) {
        subtotalCalculado = round2(subtotalCalculado + subFilaVal);
      } else if (qtyParaAlerta !== null && qtyParaAlerta > 0 && precioExcelReferencia !== null) {
        subtotalCalculado = round2(subtotalCalculado + qtyParaAlerta * precioExcelReferencia);
      }

      filasOmitidasConDatos += 1;
      advertencias.push({
        hoja,
        fila: rIdx + 1,
        codigo: '',
        mensaje: `Fila sin código con ${qtyParaAlerta ?? cantidadRaw} u.: "${descripcionExcel || 'Sin descripción'}"`,
        id: `${hoja}#${rIdx + 1}#FILA_SIN_CODIGO`,
        tipo: 'FILA_SIN_CODIGO',
        codigosDetectados: [],
        descripcionExcel,
        precioExcelReferencia,
        cantidadDetectada: qtyParaAlerta,
        cantidadRaw,
        notaDetectada: notaDetectadaRaw || modeloDetectado,
        productoExistenteIdGrupo: null,
        sugerenciasCatalogo: sugerencias.slice(0, 4)
      });
      continue;
    }

    const descNormLower = normalizarTexto(descripcionExcel);
    const modeloNormLower = normalizarTexto(modeloDetectado);
    const correccionesFila: Array<{ original: string; corregido: string; motivo: string }> = [];
    const codigosCorregidos = codigosEnFila.map((c) => {
      const corr = corregirCodigoConContexto(c, descNormLower, modeloNormLower);
      if (corr) {
        correccionesFila.push({ original: c, corregido: corr.corregido, motivo: corr.motivo });
        return corr.corregido;
      }
      return c;
    });

    // Regla oficial: ignorar fila duplicada 6545 sin precio ni cantidad
    if (
      codigosCorregidos.includes('6545') &&
      (precioCell === null || String(precioCell).trim() === '') &&
      (qtyCell === null || String(qtyCell).trim() === '' || Number(qtyCell) === 0)
    ) {
      continue;
    }

    let producto: ProductoCatalogo | undefined;
    for (const cod of codigosCorregidos) {
      const encontrado =
        catalogoPorCodigo.get(cod.toUpperCase()) ??
        CATALOGO_EXTENDIDO_BY_CODIGO.get(cod) ??
        CATALOGO_BY_CODIGO.get(cod);
      if (encontrado) {
        producto = encontrado;
        break;
      }
    }

    const codigoJoined = codigosCorregidos.join('/');
    const qtyRes = parsearCantidadEstricta(qtyCell);
    if (!qtyRes.valid) {
      advertencias.push({
        hoja,
        fila: rIdx + 1,
        codigo: codigoJoined,
        mensaje: qtyRes.motivo ?? 'Cantidad inválida',
        id: `${hoja}#${rIdx + 1}#${codigoJoined}#CANTIDAD_INVALIDA`,
        tipo: 'CANTIDAD_INVALIDA',
        codigosDetectados: codigosCorregidos,
        descripcionExcel,
        precioExcelReferencia,
        cantidadDetectada: null,
        cantidadRaw,
        notaDetectada: notaDetectadaRaw || modeloDetectado,
        productoExistenteIdGrupo: producto ? producto.idGrupo : null,
        sugerenciasCatalogo: producto
          ? []
          : sugerirProductosCercanos(codigoJoined, descripcionExcel, catalogoUnificado)
      });
      continue;
    }

    if (qtyRes.qty === 0) continue;

    // Blindaje 1: Fila oculta con cantidad > 0
    if (filasOcultas?.has(rIdx)) {
      filasOmitidasConDatos += 1;
      advertencias.push({
        hoja,
        fila: rIdx + 1,
        codigo: codigoJoined,
        mensaje: `Esta fila está oculta en el Excel pero tiene ${qtyRes.qty} u. cargadas.`,
        id: `${hoja}#${rIdx + 1}#${codigoJoined}#FILA_OCULTA`,
        tipo: 'FILA_OCULTA',
        codigosDetectados: codigosCorregidos,
        descripcionExcel,
        precioExcelReferencia,
        cantidadDetectada: qtyRes.qty,
        cantidadRaw,
        notaDetectada: notaDetectadaRaw || comentarioFilaTexto || modeloDetectado,
        productoExistenteIdGrupo: producto ? producto.idGrupo : null,
        sugerenciasCatalogo: producto
          ? []
          : sugerirProductosCercanos(codigoJoined, descripcionExcel, catalogoUnificado)
      });
      continue;
    }

    // Blindaje 2: Fila o celda tachada con cantidad > 0
    if (filaEstaTachada) {
      filasOmitidasConDatos += 1;
      advertencias.push({
        hoja,
        fila: rIdx + 1,
        codigo: codigoJoined,
        mensaje: `Esta fila aparece tachada en el Excel (${qtyRes.qty} u.).`,
        id: `${hoja}#${rIdx + 1}#${codigoJoined}#FILA_TACHADA`,
        tipo: 'FILA_TACHADA',
        codigosDetectados: codigosCorregidos,
        descripcionExcel,
        precioExcelReferencia,
        cantidadDetectada: qtyRes.qty,
        cantidadRaw,
        notaDetectada: notaDetectadaRaw || comentarioFilaTexto || modeloDetectado,
        productoExistenteIdGrupo: producto ? producto.idGrupo : null,
        sugerenciasCatalogo: producto
          ? []
          : sugerirProductosCercanos(codigoJoined, descripcionExcel, catalogoUnificado)
      });
      continue;
    }

    if (!producto) {
      const subFilaValUnk = extractNumericFromCell(subtotalCell);
      if (subFilaValUnk !== null && subFilaValUnk > 0) {
        subtotalCalculado = round2(subtotalCalculado + subFilaValUnk);
      } else if (precioExcelReferencia !== null) {
        subtotalCalculado = round2(subtotalCalculado + qtyRes.qty * precioExcelReferencia);
      }

      advertencias.push({
        hoja,
        fila: rIdx + 1,
        codigo: codigoJoined,
        mensaje: `Código desconocido (${codigoJoined}) con cantidad ${qtyRes.qty}`,
        id: `${hoja}#${rIdx + 1}#${codigoJoined}#CODIGO_DESCONOCIDO`,
        tipo: 'CODIGO_DESCONOCIDO',
        codigosDetectados: codigosCorregidos,
        descripcionExcel,
        precioExcelReferencia,
        cantidadDetectada: qtyRes.qty,
        cantidadRaw,
        notaDetectada: notaDetectadaRaw || comentarioFilaTexto || modeloDetectado,
        productoExistenteIdGrupo: null,
        sugerenciasCatalogo: sugerirProductosCercanos(codigoJoined, descripcionExcel, catalogoUnificado)
      });
      continue;
    }

    // Blindaje 5B: Control matemático por fila (cantidad * precio vs subtotal)
    const colSubtotalEfectiva = headerDetectado ? colSubtotal : 5;
    const subtotalEsFormulaSinRes = getCellFormulaSinResultado(rIdx, colSubtotalEfectiva) !== null;
    const subFilaVal = subtotalEsFormulaSinRes ? null : extractNumericFromCell(subtotalCell);
    if (
      !subtotalEsFormulaSinRes &&
      precioExcelReferencia !== null &&
      precioExcelReferencia > 0 &&
      subFilaVal !== null &&
      subFilaVal > 0
    ) {
      const esperadoNeto = round2(qtyRes.qty * precioExcelReferencia);
      const tol = Math.max(1.5, qtyRes.qty * 0.5);
      const cierraNeto = Math.abs(subFilaVal - esperadoNeto) <= tol;
      const cierraConIva = Math.abs(subFilaVal - round2(esperadoNeto * 1.21)) <= tol * 1.21;
      if (!cierraNeto && !cierraConIva) {
        const ratio = subFilaVal / precioExcelReferencia;
        const qtyImplicita =
          Math.abs(ratio - Math.round(ratio)) < 0.01 && Math.round(ratio) >= 1
            ? Math.round(ratio)
            : null;
        filasOmitidasConDatos += 1;
        advertencias.push({
          hoja,
          fila: rIdx + 1,
          codigo: codigoJoined,
          mensaje: `La fila indica ${qtyRes.qty} u. × $${precioExcelReferencia}, pero el subtotal dice $${subFilaVal}${qtyImplicita ? ` (equivalente a ${qtyImplicita} u.)` : ''}.`,
          id: `${hoja}#${rIdx + 1}#${codigoJoined}#DESCUADRE_FILA`,
          tipo: 'DESCUADRE_FILA',
          codigosDetectados: codigosCorregidos,
          descripcionExcel,
          precioExcelReferencia,
          cantidadDetectada: qtyRes.qty,
          cantidadAlternativa: qtyImplicita,
          cantidadRaw,
          notaDetectada: notaDetectadaRaw || comentarioFilaTexto || modeloDetectado,
          productoExistenteIdGrupo: producto.idGrupo,
          sugerenciasCatalogo: []
        });
        continue;
      }
    }

    // Blindaje 3: Nota base + Comentario flotante de celda
    let notaFinal = notaDetectadaRaw;
    let aplicadoComoNota = false;
    if (colT1 !== -1 && colT2 !== -1 && producto.variantes.length === 2) {
      const q1Res = parsearCantidadEstricta(row[colT1] ?? 0);
      const q2Res = parsearCantidadEstricta(row[colT2] ?? 0);
      if (q1Res.valid && q2Res.valid && q1Res.qty + q2Res.qty === qtyRes.qty) {
        notaFinal = `${producto.variantes[0]!.codigo}: ${q1Res.qty} / ${producto.variantes[1]!.codigo}: ${q2Res.qty}`;
      }
    } else if (!notaFinal && comentarioFilaTexto && producto.variantes.length >= 2) {
      notaFinal = comentarioFilaTexto;
      aplicadoComoNota = true;
    } else if (!notaFinal && producto.variantes.length >= 2) {
      if (codigosCorregidos.length === 1) {
        notaFinal = codigosCorregidos[0]!;
      } else if (modeloDetectado) {
        notaFinal = modeloDetectado;
      }
    }

    if (rowComentarios.length > 0 && comentarioFilaTexto) {
      const primerCom = rowComentarios[0]!;
      comentariosDetectados.push({
        id: `${hoja}#${rIdx + 1}#COMENTARIO`,
        hoja,
        fila: rIdx + 1,
        columna: primerCom.col,
        celdaRef: XLSX.utils.encode_cell({ r: rIdx, c: primerCom.col }),
        codigo: codigoJoined,
        productoNombre: producto.nombre,
        texto: comentarioFilaTexto,
        aplicadoComoNota
      });
    }

    // Blindaje 4: Guardián de negaciones, dudas y contradicciones (incluyendo Talle Único)
    const textoAuditable = [notaDetectadaRaw, comentarioFilaTexto].filter(Boolean).join(' · ');
    const tieneNegacionODuda = detectarNegacionODuda(textoAuditable);
    const codsSetProd = new Set(producto.variantes.map((v) => v.codigo));
    const cantidadAlternativa =
      producto.variantes.length === 1
        ? detectarCantidadContradictoriaTalleUnico(textoAuditable, qtyRes.qty, codsSetProd)
        : null;

    if (tieneNegacionODuda || cantidadAlternativa !== null) {
      filasOmitidasConDatos += 1;
      const mensajeContra = tieneNegacionODuda
        ? `La nota o comentario ("${textoAuditable}") indica cancelación o duda sobre ${producto.nombre} (${qtyRes.qty} u.).`
        : `Contradicción en ${producto.nombre}: la columna indica ${qtyRes.qty} u., pero la nota dice ${cantidadAlternativa} u. ("${textoAuditable}").`;

      advertencias.push({
        hoja,
        fila: rIdx + 1,
        codigo: codigoJoined,
        mensaje: mensajeContra,
        id: `${hoja}#${rIdx + 1}#${codigoJoined}#NOTA_CONTRADICTORIA`,
        tipo: 'NOTA_CONTRADICTORIA',
        codigosDetectados: codigosCorregidos,
        descripcionExcel,
        precioExcelReferencia,
        cantidadDetectada: qtyRes.qty,
        cantidadAlternativa,
        cantidadRaw,
        notaDetectada: textoAuditable,
        productoExistenteIdGrupo: producto.idGrupo,
        sugerenciasCatalogo: []
      });
      continue;
    }

    if (subtotalEsFormulaSinRes) {
      filasSubtotalFormulaSinResultado += 1;
    }
    if (subFilaVal !== null && subFilaVal > 0) {
      subtotalCalculado = round2(subtotalCalculado + subFilaVal);
    } else if (precioExcelReferencia !== null) {
      subtotalCalculado = round2(subtotalCalculado + qtyRes.qty * precioExcelReferencia);
    } else {
      subtotalCalculado = round2(subtotalCalculado + qtyRes.qty * producto.precioNeto);
    }

    if (
      precioExcelReferencia !== null &&
      precioExcelReferencia > 0 &&
      Math.abs(precioExcelReferencia - producto.precioNeto) > 1 &&
      !avisosPrecioMap.has(codigoJoined)
    ) {
      avisosPrecioMap.set(codigoJoined, {
        codigo: codigoJoined,
        nombreProducto: producto.nombre,
        precioPlanilla: precioExcelReferencia,
        precioVigente: producto.precioNeto
      });
    }

    lineasAcumuladas = aplicarLineaAlBorrador(lineasAcumuladas, producto, qtyRes.qty, notaFinal);

    // Blindaje 5D: Registro para detectar exactamente el mismo código repetido en varias filas
    const claveDup = codigosCorregidos
      .slice()
      .map((c) => c.trim().toUpperCase())
      .sort()
      .join('/');
    let dupEntry = aparicionesPorCodigo.get(claveDup);
    if (!dupEntry) {
      dupEntry = {
        codigo: codigoJoined,
        nombreProducto: producto.nombre,
        idGrupo: producto.idGrupo,
        apariciones: [],
        ...{}
      };
      aparicionesPorCodigo.set(claveDup, dupEntry);
    }
    dupEntry.apariciones.push({ fila: rIdx + 1, cantidad: qtyRes.qty });

    for (const cf of correccionesFila) {
      const idxCorr = correcciones.length;
      correcciones.push({
        id: `${hoja}#${rIdx + 1}#${cf.original}->${cf.corregido}`,
        hoja,
        fila: rIdx + 1,
        original: cf.original,
        corregidoA: cf.corregido,
        productoNombre: producto.nombre,
        motivo: cf.motivo,
        lineaId: producto.idGrupo
      });
      correccionesPendientesLineaId.push({ corrIdx: idxCorr, idGrupo: producto.idGrupo });
    }

    // Si resolverLinea recuperó un typo de 1 dígito en nota con códigos, transparentarlo
    if (notaFinal && producto.variantes.length >= 2) {
      const codsProd = new Set(producto.variantes.map((v) => v.codigo.toUpperCase()));
      const notaNormLocal = normalizarTexto(notaFinal);
      const lineaResuelta = lineasAcumuladas.find((l) => l.idGrupo === producto.idGrupo);
      if (lineaResuelta && lineaResuelta.estado === 'OK') {
        for (const mTok of notaNormLocal.matchAll(/\b(6\d{3})\b/g)) {
          const tok = mTok[1];
          if (!tok || codsProd.has(tok)) continue;
          const matchedVar = producto.variantes.find((v) => distanciaEdicion(tok, v.codigo.toLowerCase()) <= 1);
          if (matchedVar) {
            const idxCorr = correcciones.length;
            correcciones.push({
              id: `${hoja}#${rIdx + 1}#${tok}->${matchedVar.codigo}`,
              hoja,
              fila: rIdx + 1,
              original: tok,
              corregidoA: matchedVar.codigo,
              productoNombre: producto.nombre,
              motivo: 'Código de talle/color corregido por un dígito',
              lineaId: lineaResuelta.idLinea
            });
            correccionesPendientesLineaId.push({ corrIdx: idxCorr, idGrupo: producto.idGrupo });
          }
        }
      }
    }
  }

  for (const pending of correccionesPendientesLineaId) {
    const lineaFinal = lineasAcumuladas.find((l) => l.idGrupo === pending.idGrupo);
    if (lineaFinal && correcciones[pending.corrIdx]) {
      correcciones[pending.corrIdx]!.lineaId = lineaFinal.idLinea;
    }
  }

  // Si hubo monto de descuento al pie sin % explícito, deducir porcentaje entero cuando cierra
  if (
    descuentoDeclaradoPorc === null &&
    descuentoDeclaradoMonto !== null &&
    descuentoDeclaradoMonto > 0 &&
    subtotalDeclarado !== null &&
    subtotalDeclarado > 0
  ) {
    const impliedPct = round2((descuentoDeclaradoMonto / subtotalDeclarado) * 100);
    const pctInt = Math.round(impliedPct);
    if (pctInt > 0 && pctInt <= 50 && Math.abs(impliedPct - pctInt) <= 0.15) {
      descuentoDeclaradoPorc = pctInt;
      if (descuentoDetectado === null) descuentoDetectado = pctInt;
    }
  }

  const duplicadosEnArchivo: DuplicadoEnArchivo[] = [];
  for (const d of aparicionesPorCodigo.values()) {
    if (d.apariciones.length >= 2) {
      duplicadosEnArchivo.push({
        codigo: d.codigo,
        nombreProducto: d.nombreProducto,
        idGrupo: d.idGrupo,
        apariciones: d.apariciones,
        cantidadTotalSumada: d.apariciones.reduce((s, a) => s + a.cantidad, 0)
      });
    }
  }

  return {
    lineas: lineasAcumuladas,
    advertencias,
    avisosPrecio: Array.from(avisosPrecioMap.values()),
    correcciones,
    comentariosDetectados,
    duplicadosEnArchivo,
    erroresFormulaPie,
    formulasSinResultadoPie,
    filasSubtotalFormulaSinResultado,
    unidadesDeclaradas,
    subtotalDeclarado,
    descuentoDeclaradoPorc,
    descuentoDeclaradoMonto,
    ivaDeclaradoArchivo,
    totalConIvaDeclarado,
    subtotalCalculado: round2(subtotalCalculado),
    filasOmitidasConDatos,
    descuentoDetectado,
    headerRowIdx
  };
};

const inferirClienteDesdeNombreArchivo = (fileName: string): string => {
  const soloNombre = fileName.replace(/^.*[\\/]/, '');
  const base = soloNombre.replace(/\.[^/.]+$/, '');
  const sinParens = base.replace(/\([^)]*\)/g, ' ');
  const tokens = normalizarTexto(sinParens)
    .split(/[^a-z0-9]+/)
    .filter(Boolean);

  const NOISE = new Set([
    'nota',
    'pedido',
    'pedidos',
    'mibi',
    'orden',
    'de',
    'del',
    'la',
    'el',
    'los',
    'las',
    'final',
    'actualizado',
    'actualizada',
    'como',
    'quedaria',
    'copia',
    'version',
    'nuevo',
    'nueva',
    'verano',
    'invierno',
    'temporada',
    'enero',
    'febrero',
    'marzo',
    'abril',
    'mayo',
    'junio',
    'julio',
    'agosto',
    'septiembre',
    'octubre',
    'noviembre',
    'diciembre',
    'ene',
    'feb',
    'mar',
    'abr',
    'may',
    'jun',
    'jul',
    'ago',
    'sep',
    'sept',
    'oct',
    'nov',
    'dic',
    'lista',
    'srl',
    'sa'
  ]);

  const limpios = tokens.filter((t) => {
    if (NOISE.has(t)) return false;
    if (/^\d+$/.test(t)) return false;
    if (/^\d{2,4}[-_]\d{2,4}$/.test(t)) return false;
    return t.length >= 2;
  });

  const joined = limpios.join(' ');
  const letrasCount = (joined.match(/[a-z]/g) ?? []).length;
  if (letrasCount < 4) return '';

  return limpios.map((w) => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
};

const esValorPlaceholderCliente = (val: string): boolean => {
  const n = normalizarTexto(val);
  if (!n || n.length < 2) return true;
  return /^(?:completar|nombre|cliente|cliente \/ negocio|razon social|mibi|fecha|-+|\.+)$/.test(n);
};

export const extraerDatosCabeceraExcel = (
  rows: CellPrimitive[][],
  headerRowIdx: number,
  fileName: string,
  descuentoPie: number | null
): DatosClienteDetectados | null => {
  const limit = headerRowIdx > 0 ? Math.min(15, headerRowIdx) : Math.min(15, rows.length);
  let cliente = '';
  let cuit = '';
  let whatsapp = '';
  let transporte = '';
  let descuentoPorc: number | undefined =
    descuentoPie !== null && descuentoPie > 0 && descuentoPie < 100 ? descuentoPie : undefined;
  let fechaPedido = '';

  const hoySentinel = '0000-00-00';

  for (let rIdx = 0; rIdx < limit; rIdx++) {
    const row = rows[rIdx];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const raw = row[c];
      if (raw === null || raw === undefined) continue;
      const rawStr = limpiarTextoInvisible(String(raw));
      if (!rawStr) continue;
      const norm = normalizarTexto(rawStr);

      let label = norm;
      let inlineVal = '';
      const colonIdx = rawStr.indexOf(':');
      if (colonIdx > 0) {
        label = normalizarTexto(rawStr.slice(0, colonIdx));
        inlineVal = limpiarTextoInvisible(rawStr.slice(colonIdx + 1));
      }

      const nextNonEmptyCell = (): CellPrimitive => {
        for (let k = c + 1; k < row.length; k++) {
          const v = row[k];
          if (v !== null && v !== undefined && limpiarTextoInvisible(String(v)) !== '') {
            return v;
          }
        }
        return null;
      };

      const valCell = inlineVal ? inlineVal : nextNonEmptyCell();
      const valStr = valCell !== null && valCell !== undefined ? limpiarTextoInvisible(String(valCell)) : '';

      if (!cliente && /^(?:cliente|razon social|senor(?:es)?|nombre\s*\/\s*razon|apellido y nombre)$/.test(label)) {
        if (valStr && !esValorPlaceholderCliente(valStr)) {
          cliente = valStr;
        }
      } else if (!cuit && /^c\.?u\.?i\.?t\.?$/.test(label)) {
        if (valStr && validarCuit(valStr)) {
          cuit = normalizarCuit(valStr);
        }
      } else if (!whatsapp && /^(?:whatsapp|tel|telefono|cel|celular|tel\s*\/\s*whatsapp)$/.test(label)) {
        const digits = valStr.replace(/\D/g, '');
        if (digits.length >= 7) {
          whatsapp = valStr;
        }
      } else if (!transporte && /^(?:transporte|expreso|envio por)$/.test(label)) {
        if (valStr && !esValorPlaceholderCliente(valStr)) {
          transporte = valStr;
        }
      } else if (descuentoPorc === undefined && /^(?:descuento|bonif(?:icacion)?)$/.test(label)) {
        const num = extractNumericFromCell(valCell);
        if (num !== null && num > 0 && num < 100) {
          descuentoPorc = num;
        }
      } else if (!fechaPedido && /^(?:fecha|fecha pedido|fecha de emision)$/.test(label)) {
        const parsed = parsearFechaExcel(valCell, hoySentinel);
        if (parsed !== hoySentinel && /^\d{4}-\d{2}-\d{2}$/.test(parsed)) {
          fechaPedido = parsed;
        }
      }
    }
  }

  let origen: DatosClienteDetectados['origen'] = 'ARCHIVO';
  if (!cliente) {
    const inferido = inferirClienteDesdeNombreArchivo(fileName);
    if (inferido) {
      cliente = inferido;
      origen = 'NOMBRE_ARCHIVO';
    }
  }

  if (!cliente && !cuit && !whatsapp && !transporte && descuentoPorc === undefined && !fechaPedido) {
    return null;
  }

  return {
    ...(cliente ? { cliente } : {}),
    ...(cuit ? { cuit } : {}),
    ...(whatsapp ? { whatsapp } : {}),
    ...(transporte ? { transporte } : {}),
    ...(descuentoPorc !== undefined ? { descuentoPorc } : {}),
    ...(fechaPedido ? { fechaPedido } : {}),
    origen
  };
};

export const construirAuditoriaDesdeTabular = (
  tab: ResultadoAnalisisTabular,
  hojasDisponibles: ReadonlyArray<HojaDisponible>
): AuditoriaLectura => {
  const unidadesLeidas = tab.lineas.reduce((acc, l) => acc + l.cantidadTotal, 0);
  const unidadesEnAlertas = tab.advertencias.reduce((acc, a) => acc + (a.cantidadDetectada ?? 0), 0);
  const diferenciaUnidades =
    tab.unidadesDeclaradas !== null ? unidadesLeidas - tab.unidadesDeclaradas : null;
  const diferenciaImporte =
    tab.subtotalDeclarado !== null ? round2(tab.subtotalCalculado - tab.subtotalDeclarado) : null;

  const motivos: string[] = [];
  let estadoIntegridad: AuditoriaLectura['estadoIntegridad'] = 'EXACTO';

  if (tab.erroresFormulaPie.length > 0) {
    estadoIntegridad = 'BLOQUEADO';
    motivos.push(
      `La celda de total de la planilla contiene un error de fórmula (${tab.erroresFormulaPie[0]}).`
    );
  }
  if (tab.formulasSinResultadoPie.length > 0) {
    estadoIntegridad = estadoIntegridad === 'BLOQUEADO' ? 'BLOQUEADO' : 'REVISAR';
    motivos.push('La celda de total tiene una fórmula sin resultado guardado en el archivo.');
  }
  if (tab.filasSubtotalFormulaSinResultado > 0) {
    motivos.push(
      `Se calculó el subtotal de ${tab.filasSubtotalFormulaSinResultado} fila(s) cuya fórmula no tenía resultado guardado en el archivo.`
    );
  }
  if (diferenciaUnidades !== null && diferenciaUnidades !== 0) {
    estadoIntegridad = 'BLOQUEADO';
    motivos.push(
      `El archivo indica ${tab.unidadesDeclaradas} unidades totales, pero se leyeron ${unidadesLeidas} (diferencia de ${diferenciaUnidades > 0 ? `+${diferenciaUnidades}` : diferenciaUnidades} u.).`
    );
  }
  if (diferenciaImporte !== null && Math.abs(diferenciaImporte) > 1) {
    estadoIntegridad = 'BLOQUEADO';
    motivos.push(
      `El subtotal del archivo ($${tab.subtotalDeclarado?.toFixed(2)}) difiere de la suma de filas ($${tab.subtotalCalculado.toFixed(2)}).`
    );
  }

  // Cadena de 4 cifras al pie del Excel (Subtotal - Descuento + IVA 21% = Total Facturado)
  if (tab.subtotalDeclarado !== null && tab.totalConIvaDeclarado !== null) {
    const descMonto =
      tab.descuentoDeclaradoMonto !== null
        ? tab.descuentoDeclaradoMonto
        : tab.descuentoDeclaradoPorc !== null
          ? round2(tab.subtotalDeclarado * (tab.descuentoDeclaradoPorc / 100))
          : 0;
    const base = round2(tab.subtotalDeclarado - descMonto);
    const ivaCalculado = round2(base * 0.21);
    const ivaRef = tab.ivaDeclaradoArchivo ?? ivaCalculado;

    if (tab.ivaDeclaradoArchivo !== null && Math.abs(ivaCalculado - tab.ivaDeclaradoArchivo) > 2) {
      estadoIntegridad = 'BLOQUEADO';
      motivos.push(
        `El IVA al pie ($${tab.ivaDeclaradoArchivo.toFixed(2)}) no coincide con el 21% de la base ($${ivaCalculado.toFixed(2)}).`
      );
    } else if (Math.abs(round2(base + ivaRef) - tab.totalConIvaDeclarado) > 2) {
      estadoIntegridad = 'BLOQUEADO';
      motivos.push(
        `El total facturado al pie ($${tab.totalConIvaDeclarado.toFixed(2)}) no coincide con Base + IVA ($${round2(base + ivaRef).toFixed(2)}).`
      );
    }
  }

  if (estadoIntegridad !== 'BLOQUEADO') {
    if (
      tab.advertencias.length > 0 ||
      tab.correcciones.length > 0 ||
      tab.duplicadosEnArchivo.length > 0 ||
      hojasDisponibles.length > 1
    ) {
      estadoIntegridad = 'REVISAR';
    }
    if (tab.advertencias.length > 0) {
      motivos.push(`Hay ${tab.advertencias.length} fila(s) en el centro de alertas para resolver.`);
    }
    if (tab.correcciones.length > 0) {
      motivos.push(`Se aplicaron ${tab.correcciones.length} corrección(es) automática(s) revisables.`);
    }
    if (tab.duplicadosEnArchivo.length > 0) {
      motivos.push(
        `Se detectaron ${tab.duplicadosEnArchivo.length} código(s) repetidos en distintas filas.`
      );
    }
    if (hojasDisponibles.length > 1) {
      motivos.push(`El libro contiene ${hojasDisponibles.length} hojas con artículos; se lee solo la seleccionada.`);
    }
  }

  return {
    unidadesLeidas,
    unidadesDeclaradasArchivo: tab.unidadesDeclaradas,
    diferenciaUnidades,
    unidadesEnAlertas,
    subtotalNetoCalculado: tab.subtotalCalculado,
    totalDeclaradoArchivo: tab.subtotalDeclarado,
    diferenciaImporte,
    filasOmitidasConDatos: tab.filasOmitidasConDatos,
    estadoIntegridad,
    motivos,
    descuentoDeclaradoPorc: tab.descuentoDeclaradoPorc,
    descuentoDeclaradoMonto: tab.descuentoDeclaradoMonto,
    ivaDeclaradoArchivo: tab.ivaDeclaradoArchivo,
    totalFacturadoDeclaradoArchivo: tab.totalConIvaDeclarado,
    duplicadosEnArchivo: tab.duplicadosEnArchivo
  };
};

const detectarPlanillaGeneralEnRows = (rows: CellPrimitive[][]): boolean => {
  for (const row of rows) {
    const norms = row.map((c) => (c === null ? '' : normalizarTexto(String(c)).replace(/\s+/g, ' ').trim()));
    const tienePagoOArmado = norms.some(
      (t) =>
        t.includes('estado pago') ||
        t.includes('estado de pago') ||
        t.includes('estado armado') ||
        t.includes('estado de armado') ||
        t === 'armado'
    );
    if (!tienePagoOArmado) continue;

    let otros = 0;
    if (
      norms.some(
        (t) =>
          t.includes('n° pedido') ||
          t.includes('nro pedido') ||
          t.includes('numero pedido') ||
          t.includes('n.º pedido') ||
          t.includes('n.o pedido')
      )
    ) {
      otros++;
    }
    if (norms.some((t) => t === 'cliente' || t.startsWith('cliente ') || t.includes('cliente /'))) {
      otros++;
    }
    if (norms.some((t) => t.includes('cheque') || t.includes('saldo pendiente') || t.includes('monto abonado'))) {
      otros++;
    }
    if (norms.some((t) => t.includes('total facturado') || t.includes('total pedido'))) {
      otros++;
    }
    if (norms.some((t) => t === 'envio' || t.includes('estado envio') || t.includes('estado de envio'))) {
      otros++;
    }

    if (otros >= 2) return true;
  }
  return false;
};

export const parsearExcelLista006 = (
  buffer: ArrayBuffer,
  fileName: string,
  catalogoExtra: ReadonlyArray<ProductoCatalogo> = CATALOGO_EXTENDIDO_MIBI,
  opciones?: {
    hojaSeleccionada?: string;
    celdasTachadasPorHoja?: ReadonlyMap<string, ReadonlySet<string>>;
    comentariosThreadedPorHoja?: ReadonlyMap<string, ReadonlyMap<string, string>>;
  }
): ResultadoParseoLista006 => {
  const catalogoPorCodigo = new Map<string, ProductoCatalogo>(CATALOGO_EXTENDIDO_BY_CODIGO);
  const catalogoUnificadoMap = new Map<string, ProductoCatalogo>();
  for (const prod of CATALOGO_EXTENDIDO_MIBI) {
    catalogoUnificadoMap.set(prod.idGrupo, prod);
  }
  for (const prod of catalogoExtra) {
    catalogoUnificadoMap.set(prod.idGrupo, prod);
    for (const v of prod.variantes) {
      catalogoPorCodigo.set(v.codigo.trim().toUpperCase(), prod);
    }
  }
  const catalogoUnificado = Array.from(catalogoUnificadoMap.values());

  let wb: SheetJSWorkBook;
  try {
    wb = leerWorkbookSeguro(buffer, fileName);
  } catch {
    return {
      ok: false,
      errorFatal: 'No se pudo leer el archivo. Verificá que sea un Excel (.xlsx, .xls), LibreOffice (.ods) o CSV válido.',
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'NO_COMPATIBLE',
      auditoria: auditoriaVacia('BLOQUEADO'),
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados: null,
      comentariosDetectados: []
    };
  }

  let esLista006 = false;
  let tieneTablaEstructuradaMibi = false;
  const rowsPorHoja = new Map<string, CellPrimitive[][]>();
  const metaPorHoja = new Map<string, MetadatosHojaExcel>();

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = expandirCeldasCombinadas(ws);
    rowsPorHoja.set(sheetName, rows);
    metaPorHoja.set(
      sheetName,
      extraerMetadatosHoja(ws, opciones?.comentariosThreadedPorHoja?.get(sheetName))
    );

    if (detectarPlanillaGeneralEnRows(rows)) {
      return {
        ok: false,
        errorFatal:
          "Este archivo es una Planilla General de pedidos históricos. Para importarla, usá 'Planillas y Catálogo → Importar planilla previa'.",
        lineas: [],
        advertencias: [],
        avisosPrecio: [],
        tipoArchivo: 'PLANILLA_GENERAL_ADMIN',
        auditoria: auditoriaVacia('BLOQUEADO'),
        correccionesAutomaticas: [],
        hojasDisponibles: [],
        datosDetectados: null,
        comentariosDetectados: []
      };
    }

    let sheetHasCodigoHeader = false;
    let sheetHasQtyOrPriceHeader = false;
    for (const row of rows) {
      const rowTexts = row.map((c) => (c === null ? '' : normalizarTexto(String(c)).replace(/\s+/g, ' ').trim()));
      for (const cellText of rowTexts) {
        if (!cellText) continue;
        const txt = cellText.replace(/["°º]/g, ' ');
        if (/lista\s*(?:n\s*)?0*6\b/.test(txt)) {
          esLista006 = true;
          break;
        }
      }

      if (rowTexts.some((t) => t === 'codigos' || t === 'codigo' || t === 'codigo t1')) {
        sheetHasCodigoHeader = true;
      }
      if (
        rowTexts.some(
          (t) =>
            t.includes('cantidades pedidas') ||
            t.includes('cantidad pedida') ||
            t.includes('pedidas t1') ||
            t.includes('total pedido') ||
            t === 't1' ||
            t === 'cantidad' ||
            t === 'cant' ||
            t === 'precio c/desc'
        )
      ) {
        sheetHasQtyOrPriceHeader = true;
      }
      if (sheetHasCodigoHeader && sheetHasQtyOrPriceHeader) {
        tieneTablaEstructuradaMibi = true;
      }
    }
  }

  if (!esLista006 && !tieneTablaEstructuradaMibi) {
    return {
      ok: false,
      errorFatal:
        'El archivo subido no corresponde a una plantilla compatible con Mibi® (LISTA N° 006 o tabla con columnas Código y Cantidad).',
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'NO_COMPATIBLE',
      auditoria: auditoriaVacia('BLOQUEADO'),
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados: null,
      comentariosDetectados: []
    };
  }

  const resultadosHojas: Array<{ nombre: string; esSecundaria: boolean; rows: CellPrimitive[][]; tab: ResultadoAnalisisTabular }> = [];

  for (const sheetName of wb.SheetNames) {
    const rows = rowsPorHoja.get(sheetName);
    if (!rows) continue;
    const meta = metaPorHoja.get(sheetName);
    const sheetNorm = normalizarTexto(sheetName);
    const esSecundaria =
      sheetNorm === 'a revisar' ||
      sheetNorm === 'resumen' ||
      sheetNorm === 'revision' ||
      sheetNorm === 'control';
    const tab = analizarFilasTabulares(rows, sheetName, {
      catalogoPorCodigo,
      catalogoUnificado,
      fileName,
      filasOcultas: meta?.filasOcultas,
      celdasTachadas: opciones?.celdasTachadasPorHoja?.get(sheetName),
      erroresFormula: meta?.erroresFormula,
      formulasSinResultado: meta?.formulasSinResultado,
      comentariosCelda: meta?.comentariosCelda
    });
    resultadosHojas.push({ nombre: sheetName, esSecundaria, rows, tab });
  }

  const hayHojaPrincipalConDatos = resultadosHojas.some(
    (h) => !h.esSecundaria && (h.tab.lineas.length > 0 || h.tab.advertencias.length > 0)
  );
  const hojasCandidatas = resultadosHojas.filter(
    (h) =>
      (!h.esSecundaria || !hayHojaPrincipalConDatos) &&
      (h.tab.lineas.length > 0 || h.tab.advertencias.length > 0)
  );

  const hojaElegida =
    (opciones?.hojaSeleccionada
      ? hojasCandidatas.find((h) => h.nombre === opciones.hojaSeleccionada)
      : undefined) ??
    hojasCandidatas[0] ??
    resultadosHojas[0];

  const hojasDisponibles: HojaDisponible[] =
    hojasCandidatas.length >= 2
      ? hojasCandidatas.map((h) => ({
          nombre: h.nombre,
          lineas: h.tab.lineas.length,
          unidades: h.tab.lineas.reduce((acc, l) => acc + l.cantidadTotal, 0),
          seleccionada: hojaElegida ? h.nombre === hojaElegida.nombre : false
        }))
      : [];

  const tabFinal: ResultadoAnalisisTabular = hojaElegida
    ? hojaElegida.tab
    : {
        lineas: [],
        advertencias: [],
        avisosPrecio: [],
        correcciones: [],
        comentariosDetectados: [],
        duplicadosEnArchivo: [],
        erroresFormulaPie: [],
        formulasSinResultadoPie: [],
        filasSubtotalFormulaSinResultado: 0,
        unidadesDeclaradas: null,
        subtotalDeclarado: null,
        descuentoDeclaradoPorc: null,
        descuentoDeclaradoMonto: null,
        ivaDeclaradoArchivo: null,
        totalConIvaDeclarado: null,
        subtotalCalculado: 0,
        filasOmitidasConDatos: 0,
        descuentoDetectado: null,
        headerRowIdx: -1
      };

  const tipoArchivo = /\.csv$/i.test(fileName)
    ? 'CSV'
    : esLista006
      ? 'LISTA_006'
      : 'PLANILLA_PERSONALIZADA';

  const auditoria = construirAuditoriaDesdeTabular(tabFinal, hojasDisponibles);
  const datosDetectados = hojaElegida
    ? extraerDatosCabeceraExcel(hojaElegida.rows, tabFinal.headerRowIdx, fileName, tabFinal.descuentoDetectado)
    : extraerDatosCabeceraExcel([], -1, fileName, null);

  return {
    ok: true,
    lineas: tabFinal.lineas,
    advertencias: tabFinal.advertencias,
    avisosPrecio: tabFinal.avisosPrecio,
    tipoArchivo,
    auditoria,
    correccionesAutomaticas: tabFinal.correcciones,
    hojasDisponibles,
    datosDetectados,
    comentariosDetectados: tabFinal.comentariosDetectados
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
