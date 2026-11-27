import * as XLSX from 'xlsx';
import {
  CATALOGO_BY_CODIGO,
  CATALOGO_EXTENDIDO_BY_CODIGO,
  CATALOGO_EXTENDIDO_BY_GRUPO,
  CATALOGO_EXTENDIDO_MIBI
} from './catalogo-mibi';
import { addDaysISO, formatNowDateTime, getTodayISO, recalcularPedido, round2 } from './finanzas';
import type {
  AdvertenciaExcel,
  AvisoPrecioExcel,
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
    // "T1 (6562): 4 / T2 (6563): 8" -> "t1 6562: 4 / t2 6563: 8"
    .replace(/\(\s*([a-z0-9\-]+)\s*\)\s*(?::|=|-)?\s*(\d+)/gi, '$1: $2')
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

export const parsearExcelLista006 = (
  buffer: ArrayBuffer,
  fileName: string,
  catalogoExtra: ReadonlyArray<ProductoCatalogo> = CATALOGO_EXTENDIDO_MIBI
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
      errorFatal: 'No se pudo leer el archivo. Verificá que sea un Excel (.xlsx, .xls) o CSV válido.',
      lineas: [],
      advertencias: [],
      avisosPrecio: []
    };
  }

  let esLista006 = false;
  let tieneTablaEstructuradaMibi = false;
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = expandirCeldasCombinadas(ws);
    let sheetHasCodigoHeader = false;
    let sheetHasQtyOrPriceHeader = false;
    for (const row of rows) {
      const rowTexts = row.map((c) => (c === null ? '' : normalizarTexto(String(c))));
      for (const cellText of rowTexts) {
        if (!cellText) continue;
        const txt = cellText.replace(/["°º]/g, ' ');
        if (/lista\s*(?:n\s*)?0*6\b/.test(txt)) {
          esLista006 = true;
          break;
        }
      }
      if (esLista006) break;

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
    if (esLista006) break;
  }

  if (!esLista006 && !tieneTablaEstructuradaMibi) {
    return {
      ok: false,
      errorFatal:
        'El archivo subido no corresponde a una plantilla compatible con Mibi® (LISTA N° 006 o tabla con columnas Código y Cantidad).',
      lineas: [],
      advertencias: [],
      avisosPrecio: []
    };
  }

  const advertencias: AdvertenciaExcel[] = [];
  const avisosPrecioMap = new Map<string, AvisoPrecioExcel>();
  let lineasAcumuladas: LineaPedido[] = [];

  for (const sheetName of wb.SheetNames) {
    const sheetNorm = normalizarTexto(sheetName);
    if ((sheetNorm === 'a revisar' || sheetNorm === 'resumen') && lineasAcumuladas.length > 0) {
      continue;
    }

    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = expandirCeldasCombinadas(ws);

    let colCodeA = 0;
    let colCodeEnd = 1; // rango inclusivo [colCodeA..colCodeEnd] para soportar 1 o más columnas de códigos antes de la descripción
    let colDesc = 2;
    let colPrecio = 3;
    let colQty = 4;
    let colTotalPed = -1;
    let colT1 = -1;
    let colT2 = -1;
    let colNota = -1;
    let colModelo = -1;
    let headerDetectado = false;
    let hadNotaHeader = false;

    for (let rIdx = 0; rIdx < rows.length; rIdx++) {
      const row = rows[rIdx];
      if (!row) continue;
      const rowTexts = row.map((c) => (c === null ? '' : normalizarTexto(String(c))));

      const idxCantPed = rowTexts.findIndex(
        (t) => t.includes('cantidades pedidas') || t.includes('cantidad pedida')
      );
      const idxT1 = rowTexts.findIndex((t) => t.includes('pedidas t1') || t === 't1' || t === 'talle 1');
      const idxT2 = rowTexts.findIndex((t) => t.includes('pedidas t2') || t === 't2' || t === 'talle 2');
      const idxTotalPed = rowTexts.findIndex((t) => t.includes('total pedido'));
      const idxCantFallback = rowTexts.findIndex((t) => t === 'cantidad' || t === 'cant');
      const idxTalleColor = rowTexts.findIndex(
        (t) => t.includes('talles/colores') || t.includes('talles / colores') || t.includes('talles/ colores')
      );
      const idxModelo = rowTexts.findIndex((t) => t.includes('modelo') || t === 'talle');
      const idxDetalleFallback = rowTexts.findIndex((t) => t.includes('detalle') && !t.includes('producto'));
      const idxCodigos = rowTexts.findIndex((t) => t === 'codigos' || t === 'codigo' || t === 'codigo t1');
      const idxCodigoT2 = rowTexts.findIndex((t) => t === 'codigo t2');
      const idxPrecio = rowTexts.findIndex((t) => t === 'precio' || t.includes('precio neto'));
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
        if (idxDesc !== -1) {
          colDesc = idxDesc;
          if (colCodeEnd >= idxDesc && idxDesc > colCodeA) {
            colCodeEnd = idxDesc - 1;
          }
        }
        if (idxPrecio !== -1) colPrecio = idxPrecio;
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
        } else if (idxDetalleFallback !== -1) {
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

      const codigosEnFila: string[] = [];
      let esFilaTotal = false;
      for (let cIdx = Math.max(0, colCodeA); cIdx <= Math.min(row.length - 1, colCodeEnd); cIdx++) {
        const rawCell = row[cIdx];
        if (rawCell === null || rawCell === undefined) continue;
        const cleanCell = limpiarTextoInvisible(String(rawCell)).toUpperCase();
        if (!cleanCell) continue;
        if (cIdx === colCodeA && /^(?:TOTAL|SUBTOTAL)\b/.test(cleanCell)) {
          esFilaTotal = true;
          break;
        }
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
      if (esFilaTotal || codigosEnFila.length === 0) continue;

      const descCell = headerDetectado ? (row[colDesc] ?? null) : (row[2] ?? null);
      const precioCell = headerDetectado ? (colPrecio >= 0 ? (row[colPrecio] ?? null) : null) : (row[3] ?? null);
      const notaCell = headerDetectado ? (colNota >= 0 ? (row[colNota] ?? null) : null) : (row[6] ?? null);
      const modeloCell = headerDetectado && colModelo >= 0 ? (row[colModelo] ?? null) : null;

      // Cantidad: en plantillas con T1/T2 sin columna Total Pedido, sumar T1 + T2
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
      if (/^(?:total|subtotal)\b/i.test(descripcionExcel)) continue;

      const notaDetectadaRaw =
        notaCell !== null && notaCell !== undefined ? limpiarTextoInvisible(String(notaCell)) : '';
      const modeloDetectado =
        modeloCell !== null && modeloCell !== undefined ? limpiarTextoInvisible(String(modeloCell)) : '';
      const cantidadRaw =
        qtyCell !== null && qtyCell !== undefined ? limpiarTextoInvisible(String(qtyCell)) : '';

      // Typo conocido en planilla Mazzaro: 6570 Chaleco Bug -> 6540
      const descNormLower = normalizarTexto(descripcionExcel);
      const codigosCorregidos = codigosEnFila.map((c) =>
        c === '6570' && descNormLower.includes('chaleco') && (descNormLower.includes('bug') || normalizarTexto(modeloDetectado).includes('bug'))
          ? '6540'
          : c
      );

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
          hoja: sheetName,
          fila: rIdx + 1,
          codigo: codigoJoined,
          mensaje: qtyRes.motivo ?? 'Cantidad inválida',
          id: `${sheetName}#${rIdx + 1}#${codigoJoined}#CANTIDAD_INVALIDA`,
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

      if (!producto) {
        advertencias.push({
          hoja: sheetName,
          fila: rIdx + 1,
          codigo: codigoJoined,
          mensaje: `Código desconocido (${codigoJoined}) con cantidad ${qtyRes.qty}`,
          id: `${sheetName}#${rIdx + 1}#${codigoJoined}#CODIGO_DESCONOCIDO`,
          tipo: 'CODIGO_DESCONOCIDO',
          codigosDetectados: codigosCorregidos,
          descripcionExcel,
          precioExcelReferencia,
          cantidadDetectada: qtyRes.qty,
          cantidadRaw,
          notaDetectada: notaDetectadaRaw || modeloDetectado,
          productoExistenteIdGrupo: null,
          sugerenciasCatalogo: sugerirProductosCercanos(codigoJoined, descripcionExcel, catalogoUnificado)
        });
        continue;
      }

      // Construir nota final con prioridad: columnas T1/T2 separadas -> nota explícita -> código único / modelo
      let notaFinal = notaDetectadaRaw;
      if (colT1 !== -1 && colT2 !== -1 && producto.variantes.length === 2) {
        const q1Res = parsearCantidadEstricta(row[colT1] ?? 0);
        const q2Res = parsearCantidadEstricta(row[colT2] ?? 0);
        if (q1Res.valid && q2Res.valid && q1Res.qty + q2Res.qty === qtyRes.qty) {
          notaFinal = `${producto.variantes[0]!.codigo}: ${q1Res.qty} / ${producto.variantes[1]!.codigo}: ${q2Res.qty}`;
        }
      } else if (!notaFinal && producto.variantes.length >= 2) {
        if (codigosCorregidos.length === 1) {
          notaFinal = codigosCorregidos[0]!;
        } else if (modeloDetectado) {
          notaFinal = modeloDetectado;
        }
      }

      // Auditoría de precio informativa sin confiar nunca en el precio de la planilla
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
    }
  }

  return {
    ok: true,
    lineas: lineasAcumuladas,
    advertencias,
    avisosPrecio: Array.from(avisosPrecioMap.values())
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
