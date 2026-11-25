import 'server-only';

import { randomUUID } from 'node:crypto';
import { and, desc, eq, inArray, isNull, sql } from 'drizzle-orm';
import { CATALOGO_BY_GRUPO } from '@/lib/catalogo-mibi';
import { getDb, isLocalFallbackMode, readLocalStore, writeLocalStore } from '@/lib/db';
import {
  pedidoAsignaciones,
  pedidoHistorial,
  pedidoLineas,
  pedidos,
  productosCustom,
  usuarios
} from '@/lib/db/schema';
import { normalizarTexto } from '@/lib/excel-parser';
import {
  addDaysISO,
  formatNowDateTime,
  formatShortDate,
  getTodayISO,
  recalcularPedido,
  round2
} from '@/lib/finanzas';
import {
  isDepositoRole,
  rolUsuarioToLabel,
  type EstadoArmado,
  type EstadoEnvio,
  type HistorialEvento,
  type LineaPedido,
  type MedioPago,
  type Pedido,
  type PedidoAsignacion,
  type PedidoBase,
  type PrioridadAsignacion,
  type ProductoCustomRecord,
  type RolUsuario,
  type SessionUser
} from '@/lib/types';
import { AuthError, requireAdmin, requireSession } from './auth';
import { listarProductosCustomDal } from './catalogo';

export interface CrearPedidoInput {
  cliente: string;
  vendedor: string;
  whatsapp: string;
  fechaPedido: string;
  despachoPrevisto: string;
  observaciones: string;
  descuentoPorc: number;
  medioPago: MedioPago;
  montoAbonado: number;
  montoCheque: number;
  fechaRecepcionCheque: string;
  fechaAcreditacionCheque: string;
  detalleCheque: string;
  lineas: LineaPedido[];
}

export interface PatchPedidoInput {
  cliente?: string;
  vendedor?: string;
  whatsapp?: string;
  despachoPrevisto?: string;
  descuentoPorc?: number;
  medioPago?: MedioPago;
  montoAbonado?: number;
  montoCheque?: number;
  chequeAcreditado?: boolean;
  fechaRecepcionCheque?: string;
  fechaAcreditacionCheque?: string;
  detalleCheque?: string;
  estadoArmado?: EstadoArmado;
  estadoEnvio?: EstadoEnvio;
  fechaDespacho?: string;
  fechaEntrega?: string;
  transporte?: string;
  nroSeguimiento?: string;
  observaciones?: string;
  lineas?: LineaPedido[];
}

export type BatchActionInput =
  | { type: 'armado'; estadoArmado: 'En armado' | 'Listo' }
  | { type: 'envio'; estadoEnvio: 'Despachado' | 'Entregado'; transporte?: string }
  | { type: 'pago'; estado: 'Pagado' | 'Falta pago' }
  | { type: 'vendedor'; vendedor: string };

export interface EnviarCuentasInput {
  pedidoIds: string[];
  usuarioIds: string[];
  prioridad: PrioridadAsignacion;
  notaEnvio: string;
  forzarSinPago?: boolean;
}

export interface EnviarCuentasResult {
  ok: boolean;
  enviados: number;
  omitidosTalle: string[];
  omitidosSinLineas: string[];
  requiereConfirmacionSinPago?: boolean;
  pedidosSinPago?: string[];
  error?: string;
}

export interface SyncStatusResult {
  maxUpdatedAt: number;
  activeCount: number;
  fingerprint: string;
}

const ADMIN_ONLY_PATCH_KEYS: ReadonlyArray<keyof PatchPedidoInput> = [
  'vendedor',
  'descuentoPorc',
  'medioPago',
  'montoAbonado',
  'montoCheque',
  'chequeAcreditado',
  'fechaRecepcionCheque',
  'fechaAcreditacionCheque',
  'detalleCheque'
];

const DEPOSITO_ALLOWED_PATCH_KEYS = new Set<keyof PatchPedidoInput>([
  'estadoArmado',
  'estadoEnvio',
  'fechaDespacho',
  'fechaEntrega',
  'transporte',
  'nroSeguimiento',
  'observaciones'
]);

const toNumericString = (val: number): string => val.toFixed(2);
const fromNumericString = (val: string | null | undefined): number => {
  if (!val) return 0;
  const n = Number(val);
  return Number.isFinite(n) ? n : 0;
};

const sanitizarLineasContraCatalogo = (
  lineas: ReadonlyArray<LineaPedido>,
  customCatalogo: ReadonlyArray<ProductoCustomRecord>
): LineaPedido[] => {
  const customByGrupo = new Map<string, ProductoCustomRecord>();
  for (const c of customCatalogo) {
    customByGrupo.set(c.idGrupo, c);
  }

  return lineas.map((l) => {
    const prodOficial = CATALOGO_BY_GRUPO.get(l.idGrupo);
    const prodCustom = customByGrupo.get(l.idGrupo);
    const esAdhoc = l.idGrupo.startsWith('ADHOC-');

    if (!prodOficial && !prodCustom && !esAdhoc) {
      throw new Error(`Producto desconocido en catálogo: ${l.idGrupo}`);
    }

    const sumaDesglose = l.desglose.reduce((acc, d) => acc + Math.max(0, Math.trunc(d.cantidad)), 0);
    const desgloseValido = l.estado === 'OK' && !l.requiereRevision && sumaDesglose === l.cantidadTotal;

    if (prodOficial) {
      return {
        ...l,
        nombreProducto: prodOficial.nombre,
        rubro: prodOficial.rubro,
        precioUnitarioNeto: prodOficial.precioNeto,
        cantidadTotal: Math.max(1, Math.trunc(l.cantidadTotal)),
        estado: desgloseValido ? 'OK' : 'REQUIERE_REVISION',
        requiereRevision: !desgloseValido
      };
    }

    if (prodCustom) {
      return {
        ...l,
        nombreProducto: prodCustom.nombre,
        rubro: prodCustom.rubro,
        precioUnitarioNeto: prodCustom.precioNeto,
        cantidadTotal: Math.max(1, Math.trunc(l.cantidadTotal)),
        estado: desgloseValido ? 'OK' : 'REQUIERE_REVISION',
        requiereRevision: !desgloseValido
      };
    }

    // Producto fuera de planilla ad-hoc (solo para este pedido)
    const nombre = l.nombreProducto.trim();
    if (!nombre || l.desglose.length === 0) {
      throw new Error('Producto fuera de planilla inválido.');
    }
    return {
      ...l,
      nombreProducto: nombre,
      rubro: l.rubro.trim().toUpperCase() || 'OTROS',
      precioUnitarioNeto: round2(Math.max(0, l.precioUnitarioNeto)),
      cantidadTotal: Math.max(1, Math.trunc(l.cantidadTotal)),
      estado: desgloseValido ? 'OK' : 'REQUIERE_REVISION',
      requiereRevision: !desgloseValido
    };
  });
};

const sortPedidos = (list: Pedido[]): Pedido[] =>
  list.sort((a, b) => {
    const diff = b.createdAt - a.createdAt;
    if (diff !== 0) return diff;
    return (parseInt(b.numeroPedido, 10) || 0) - (parseInt(a.numeroPedido, 10) || 0);
  });

const FINANCIAL_HISTORY_RE = /\$|pagad|cobrad|cheque|monto|descuento|acredit|facturado|sin pagar|falta pago/i;

export const sanitizePedidoForRole = (pedido: Pedido, rol: RolUsuario, viewerUserId?: string): Pedido => {
  if (!isDepositoRole(rol)) return pedido;

  const observacionesLimpias = pedido.observaciones
    .split('·')
    .map((seg) => seg.trim())
    .filter((seg) => seg && !/\$|cheque acredit|cobro confirm/i.test(seg))
    .join(' · ');

  const historialLimpio = pedido.historial
    .filter((h) => !FINANCIAL_HISTORY_RE.test(h.accion))
    .map((h) => ({ ...h, accion: h.accion.replace(/\$\s*[\d.,]+/g, '[oculto]') }));

  const asignacionesVisibles = viewerUserId
    ? pedido.asignaciones.filter((a) => a.usuarioId === viewerUserId)
    : pedido.asignaciones.filter((a) => isDepositoRole(a.usuarioRol));

  return {
    ...pedido,
    lineas: pedido.lineas.map((l) => ({
      ...l,
      precioUnitarioNeto: 0
    })),
    descuentoPorc: 0,
    subtotalNeto: 0,
    montoDescuento: 0,
    baseImponible: 0,
    iva21: 0,
    totalFacturado: 0,
    totalFacturadoOverride: undefined,
    medioPago: 'Sin pagar',
    montoAbonado: 0,
    montoCheque: 0,
    chequeAcreditado: false,
    fechaRecepcionCheque: '',
    fechaAcreditacionCheque: '',
    detalleCheque: '',
    saldoPendiente: 0,
    deudaSinRespaldo: 0,
    estadoPago: 'Falta pago',
    observaciones: observacionesLimpias,
    historial: historialLimpio,
    asignaciones: asignacionesVisibles
  };
};

const isPedidoVisibleForSession = (pedido: Pedido, session: SessionUser): boolean => {
  if (session.rol === 'ADMIN') return true;
  const asignado = pedido.asignaciones.some((a) => a.usuarioId === session.userId);
  if (isDepositoRole(session.rol)) {
    return asignado;
  }
  // VENTAS o VENTAS_DEPOSITO
  if (pedido.createdBy && pedido.createdBy === session.userId) return true;
  if (asignado) return true;
  if (!pedido.createdBy && normalizarTexto(pedido.vendedor) === normalizarTexto(session.nombre)) {
    return true;
  }
  return false;
};

const mapDbRowsToPedidos = (
  pedidoRows: Array<typeof pedidos.$inferSelect>,
  lineasRows: Array<typeof pedidoLineas.$inferSelect>,
  historialRows: Array<typeof pedidoHistorial.$inferSelect>,
  asignacionesRows: Array<{
    asig: typeof pedidoAsignaciones.$inferSelect;
    usuarioNombre: string | null;
    usuarioRol: RolUsuario | null;
  }>
): Pedido[] => {
  const lineasByPedido = new Map<string, LineaPedido[]>();
  for (const lr of lineasRows) {
    const arr = lineasByPedido.get(lr.pedidoId) ?? [];
    arr.push({
      idLinea: lr.idLinea,
      idGrupo: lr.idGrupo,
      nombreProducto: lr.nombreProducto,
      rubro: lr.rubro,
      precioUnitarioNeto: fromNumericString(lr.precioUnitarioNeto),
      cantidadTotal: lr.cantidadTotal,
      notaCliente: lr.notaCliente,
      estado: lr.estado,
      requiereRevision: lr.requiereRevision,
      motivoRevision: lr.motivoRevision ?? undefined,
      desglose: lr.desgloseJson
    });
    lineasByPedido.set(lr.pedidoId, arr);
  }

  const histByPedido = new Map<string, HistorialEvento[]>();
  for (const hr of historialRows) {
    const arr = histByPedido.get(hr.pedidoId) ?? [];
    const dd = String(hr.createdAt.getDate()).padStart(2, '0');
    const mm = String(hr.createdAt.getMonth() + 1).padStart(2, '0');
    const hh = String(hr.createdAt.getHours()).padStart(2, '0');
    const min = String(hr.createdAt.getMinutes()).padStart(2, '0');
    arr.push({
      fechaHora: `${dd}/${mm} ${hh}:${min}`,
      usuario: hr.usuarioNombre,
      rol:
        hr.rol === 'Administración'
          ? 'Administración'
          : hr.rol === 'Depósito'
            ? 'Depósito'
            : hr.rol === 'Ventas'
              ? 'Ventas'
              : 'Ventas / Depósito',
      accion: hr.accion
    });
    histByPedido.set(hr.pedidoId, arr);
  }

  const asigByPedido = new Map<string, PedidoAsignacion[]>();
  for (const row of asignacionesRows) {
    const a = row.asig;
    const arr = asigByPedido.get(a.pedidoId) ?? [];
    arr.push({
      id: a.id,
      pedidoId: a.pedidoId,
      usuarioId: a.usuarioId,
      usuarioNombre: row.usuarioNombre ?? 'Cuenta',
      usuarioRol: row.usuarioRol ?? 'DEPOSITO',
      enviadoPor: a.enviadoPor ?? null,
      enviadoPorNombre: a.enviadoPorNombre,
      notaEnvio: a.notaEnvio,
      prioridad: a.prioridad,
      vistoAt: a.vistoAt ? a.vistoAt.getTime() : null,
      requiereRelectura: a.requiereRelectura,
      createdAt: a.createdAt.getTime(),
      updatedAt: a.updatedAt.getTime()
    });
    asigByPedido.set(a.pedidoId, arr);
  }

  const result = pedidoRows.map((pr) => {
    const override =
      pr.totalFacturadoOverride !== null ? fromNumericString(pr.totalFacturadoOverride) : undefined;
    const base: PedidoBase = {
      id: pr.id,
      numeroPedido: pr.numeroPedido,
      cliente: pr.cliente,
      vendedor: pr.vendedor,
      whatsapp: pr.whatsapp,
      fechaPedido: pr.fechaPedido,
      despachoPrevisto: pr.despachoPrevisto ?? '',
      fechaDespacho: pr.fechaDespacho ?? undefined,
      fechaEntrega: pr.fechaEntrega ?? undefined,
      lineas: lineasByPedido.get(pr.id) ?? [],
      descuentoPorc: fromNumericString(pr.descuentoPorc),
      medioPago: pr.medioPago,
      montoAbonado: fromNumericString(pr.montoAbonado),
      montoCheque: fromNumericString(pr.montoCheque),
      chequeAcreditado: pr.chequeAcreditado,
      fechaRecepcionCheque: pr.fechaRecepcionCheque ?? '',
      fechaAcreditacionCheque: pr.fechaAcreditacionCheque ?? '',
      detalleCheque: pr.detalleCheque,
      estadoArmado: pr.estadoArmado,
      estadoEnvio: pr.estadoEnvio,
      transporte: pr.transporte,
      nroSeguimiento: pr.nroSeguimiento,
      observaciones: pr.observaciones,
      createdBy: pr.createdBy ?? null,
      incidenciaDeposito: pr.incidenciaDeposito,
      incidenciaAt: pr.incidenciaAt ? pr.incidenciaAt.getTime() : null,
      incidenciaPor: pr.incidenciaPor,
      asignaciones: asigByPedido.get(pr.id) ?? [],
      version: pr.version,
      deletedAt: pr.deletedAt ? pr.deletedAt.toISOString() : null,
      createdAt: pr.createdAt.getTime(),
      updatedAt: pr.updatedAt.getTime(),
      historial: histByPedido.get(pr.id) ?? [],
      detalleManualImportado: pr.detalleManualImportado ?? undefined,
      totalFacturadoOverride: override
    };
    return recalcularPedido(base);
  });

  return sortPedidos(result);
};

const attachLocalAsignaciones = (
  pedidosList: ReadonlyArray<Pedido>,
  allAsig: ReadonlyArray<PedidoAsignacion>
): Pedido[] => {
  const byPedido = new Map<string, PedidoAsignacion[]>();
  for (const a of allAsig) {
    const arr = byPedido.get(a.pedidoId) ?? [];
    arr.push(a);
    byPedido.set(a.pedidoId, arr);
  }
  return pedidosList.map((p) =>
    recalcularPedido({
      ...p,
      asignaciones: byPedido.get(p.id) ?? []
    })
  );
};

export const listarPedidosDal = async (): Promise<Pedido[]> => {
  const session = await requireSession();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const activos = attachLocalAsignaciones(
      store.pedidos.filter((p) => !p.deletedAt),
      store.pedidoAsignaciones
    );
    const visibles = activos
      .filter((p) => isPedidoVisibleForSession(p, session))
      .map((p) => sanitizePedidoForRole(p, session.rol, session.userId));
    return sortPedidos(visibles);
  }

  const db = getDb();
  const pedidoRows = await db
    .select()
    .from(pedidos)
    .where(isNull(pedidos.deletedAt))
    .orderBy(desc(pedidos.createdAt));

  if (pedidoRows.length === 0) return [];
  const ids = pedidoRows.map((p) => p.id);

  const [lineasRows, historialRows, asignacionesRows] = await Promise.all([
    db.select().from(pedidoLineas).where(inArray(pedidoLineas.pedidoId, ids)),
    db
      .select()
      .from(pedidoHistorial)
      .where(inArray(pedidoHistorial.pedidoId, ids))
      .orderBy(desc(pedidoHistorial.createdAt)),
    db
      .select({
        asig: pedidoAsignaciones,
        usuarioNombre: usuarios.nombre,
        usuarioRol: usuarios.rol
      })
      .from(pedidoAsignaciones)
      .leftJoin(usuarios, eq(pedidoAsignaciones.usuarioId, usuarios.id))
      .where(inArray(pedidoAsignaciones.pedidoId, ids))
  ]);

  const mapped = mapDbRowsToPedidos(pedidoRows, lineasRows, historialRows, asignacionesRows);
  return mapped
    .filter((p) => isPedidoVisibleForSession(p, session))
    .map((p) => sanitizePedidoForRole(p, session.rol, session.userId));
};

export const obtenerPedidoPorIdDal = async (id: string): Promise<Pedido | null> => {
  const session = await requireSession();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const found = store.pedidos.find((p) => p.id === id && !p.deletedAt);
    if (!found) return null;
    const withAsig = attachLocalAsignaciones([found], store.pedidoAsignaciones)[0]!;
    if (!isPedidoVisibleForSession(withAsig, session)) {
      throw new AuthError('No tenés acceso a este pedido.', 403);
    }
    return sanitizePedidoForRole(withAsig, session.rol, session.userId);
  }

  const db = getDb();
  const pedidoRows = await db
    .select()
    .from(pedidos)
    .where(and(eq(pedidos.id, id), isNull(pedidos.deletedAt)))
    .limit(1);

  if (pedidoRows.length === 0) return null;
  const [lineasRows, historialRows, asignacionesRows] = await Promise.all([
    db.select().from(pedidoLineas).where(eq(pedidoLineas.pedidoId, id)),
    db
      .select()
      .from(pedidoHistorial)
      .where(eq(pedidoHistorial.pedidoId, id))
      .orderBy(desc(pedidoHistorial.createdAt)),
    db
      .select({
        asig: pedidoAsignaciones,
        usuarioNombre: usuarios.nombre,
        usuarioRol: usuarios.rol
      })
      .from(pedidoAsignaciones)
      .leftJoin(usuarios, eq(pedidoAsignaciones.usuarioId, usuarios.id))
      .where(eq(pedidoAsignaciones.pedidoId, id))
  ]);

  const ped = mapDbRowsToPedidos(pedidoRows, lineasRows, historialRows, asignacionesRows)[0];
  if (!ped) return null;
  if (!isPedidoVisibleForSession(ped, session)) {
    throw new AuthError('No tenés acceso a este pedido.', 403);
  }
  return sanitizePedidoForRole(ped, session.rol, session.userId);
};

export const obtenerPedidosPorIdsDal = async (ids: ReadonlyArray<string>): Promise<Pedido[]> => {
  const session = await requireSession();
  if (ids.length === 0) return [];

  if (isLocalFallbackMode()) {
    const idSet = new Set(ids);
    const store = readLocalStore();
    const matched = attachLocalAsignaciones(
      store.pedidos.filter((p) => !p.deletedAt && idSet.has(p.id)),
      store.pedidoAsignaciones
    );
    return sortPedidos(
      matched
        .filter((p) => isPedidoVisibleForSession(p, session))
        .map((p) => sanitizePedidoForRole(p, session.rol, session.userId))
    );
  }

  const db = getDb();
  const pedidoRows = await db
    .select()
    .from(pedidos)
    .where(and(inArray(pedidos.id, [...ids]), isNull(pedidos.deletedAt)))
    .orderBy(desc(pedidos.createdAt));
  if (pedidoRows.length === 0) return [];

  const foundIds = pedidoRows.map((p) => p.id);
  const [lineasRows, historialRows, asignacionesRows] = await Promise.all([
    db.select().from(pedidoLineas).where(inArray(pedidoLineas.pedidoId, foundIds)),
    db
      .select()
      .from(pedidoHistorial)
      .where(inArray(pedidoHistorial.pedidoId, foundIds))
      .orderBy(desc(pedidoHistorial.createdAt)),
    db
      .select({
        asig: pedidoAsignaciones,
        usuarioNombre: usuarios.nombre,
        usuarioRol: usuarios.rol
      })
      .from(pedidoAsignaciones)
      .leftJoin(usuarios, eq(pedidoAsignaciones.usuarioId, usuarios.id))
      .where(inArray(pedidoAsignaciones.pedidoId, foundIds))
  ]);

  return mapDbRowsToPedidos(pedidoRows, lineasRows, historialRows, asignacionesRows)
    .filter((p) => isPedidoVisibleForSession(p, session))
    .map((p) => sanitizePedidoForRole(p, session.rol, session.userId));
};

export const getSyncStatusDal = async (): Promise<SyncStatusResult> => {
  const session = await requireSession();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const visibles = attachLocalAsignaciones(
      store.pedidos.filter((p) => !p.deletedAt),
      store.pedidoAsignaciones
    ).filter((p) => isPedidoVisibleForSession(p, session));

    const maxPed = visibles.reduce((acc, p) => Math.max(acc, p.updatedAt), 0);
    const maxAsig = visibles.reduce(
      (acc, p) => Math.max(acc, ...p.asignaciones.map((a) => a.updatedAt)),
      0
    );
    const maxCat = store.productosCustom.reduce((acc, c) => Math.max(acc, c.updatedAt), 0);
    const asigCount = visibles.reduce((acc, p) => acc + p.asignaciones.length, 0);
    const maxUpdatedAt = Math.max(maxPed, maxAsig, maxCat);
    const fingerprint = `${visibles.length}:${asigCount}:${store.productosCustom.length}:${maxUpdatedAt}`;
    return { maxUpdatedAt, activeCount: visibles.length, fingerprint };
  }

  const db = getDb();
  const [pedRows, asigRows, catRows] = await Promise.all([
    db
      .select({
        maxUpdated: sql<Date | null>`max(${pedidos.updatedAt})`,
        count: sql<number>`count(*)::int`
      })
      .from(pedidos)
      .where(isNull(pedidos.deletedAt)),
    session.rol === 'ADMIN'
      ? db
          .select({
            maxUpdated: sql<Date | null>`max(${pedidoAsignaciones.updatedAt})`,
            count: sql<number>`count(*)::int`
          })
          .from(pedidoAsignaciones)
      : db
          .select({
            maxUpdated: sql<Date | null>`max(${pedidoAsignaciones.updatedAt})`,
            count: sql<number>`count(*)::int`
          })
          .from(pedidoAsignaciones)
          .where(eq(pedidoAsignaciones.usuarioId, session.userId)),
    db
      .select({
        maxUpdated: sql<Date | null>`max(${productosCustom.updatedAt})`,
        count: sql<number>`count(*)::int`
      })
      .from(productosCustom)
  ]);

  const maxPed = pedRows[0]?.maxUpdated ? new Date(pedRows[0].maxUpdated).getTime() : 0;
  const maxAsig = asigRows[0]?.maxUpdated ? new Date(asigRows[0].maxUpdated).getTime() : 0;
  const maxCat = catRows[0]?.maxUpdated ? new Date(catRows[0].maxUpdated).getTime() : 0;
  const maxUpdatedAt = Math.max(maxPed, maxAsig, maxCat);
  const pCount = pedRows[0]?.count ?? 0;
  const aCount = asigRows[0]?.count ?? 0;
  const cCount = catRows[0]?.count ?? 0;
  return {
    maxUpdatedAt,
    activeCount: pCount,
    fingerprint: `${pCount}:${aCount}:${cCount}:${maxUpdatedAt}`
  };
};

export const crearPedidoDal = async (input: CrearPedidoInput): Promise<Pedido> => {
  const session = await requireSession();
  if (isDepositoRole(session.rol)) {
    throw new AuthError('Las cuentas de Depósito no pueden crear pedidos.', 403);
  }
  const esAdmin = session.rol === 'ADMIN';

  if (!esAdmin) {
    if (
      input.medioPago !== 'Sin pagar' ||
      input.montoAbonado > 0 ||
      input.montoCheque > 0 ||
      input.descuentoPorc > 0
    ) {
      throw new AuthError('Solo Administración puede registrar pagos, cheques o descuentos.', 403);
    }
  }

  const customCatalogo = await listarProductosCustomDal(false);
  const lineasSeguras = sanitizarLineasContraCatalogo(input.lineas, customCatalogo);
  if (lineasSeguras.length === 0) {
    throw new Error('El pedido debe tener al menos un artículo.');
  }

  const hoy = getTodayISO();
  const ahora = Date.now();
  const rolLabel = rolUsuarioToLabel(session.rol);

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const maxActual = store.pedidos.reduce((acc, p) => {
      const n = parseInt(p.numeroPedido.replace(/\D/g, ''), 10);
      return Number.isFinite(n) && n > acc ? n : acc;
    }, store.nextNumeroSeq);
    const nextSeq = maxActual + 1;
    const numeroPedido = String(nextSeq).padStart(4, '0');
    store.nextNumeroSeq = nextSeq;

    const nuevo = recalcularPedido({
      id: randomUUID(),
      numeroPedido,
      cliente: input.cliente.trim(),
      vendedor: input.vendedor,
      whatsapp: input.whatsapp.trim(),
      fechaPedido: input.fechaPedido || hoy,
      despachoPrevisto: input.despachoPrevisto || addDaysISO(hoy, 5),
      lineas: lineasSeguras,
      descuentoPorc: esAdmin ? input.descuentoPorc : 0,
      medioPago: esAdmin ? input.medioPago : 'Sin pagar',
      montoAbonado: esAdmin ? input.montoAbonado : 0,
      montoCheque: esAdmin ? input.montoCheque : 0,
      chequeAcreditado: false,
      fechaRecepcionCheque: esAdmin && input.medioPago === 'Cheque' ? input.fechaRecepcionCheque : '',
      fechaAcreditacionCheque: esAdmin && input.medioPago === 'Cheque' ? input.fechaAcreditacionCheque : '',
      detalleCheque: esAdmin && input.medioPago === 'Cheque' ? input.detalleCheque.trim() : '',
      estadoArmado: 'Por armar',
      estadoEnvio: 'Pendiente',
      transporte: '',
      nroSeguimiento: '',
      observaciones: input.observaciones.trim(),
      createdBy: session.userId,
      incidenciaDeposito: '',
      incidenciaAt: null,
      incidenciaPor: '',
      asignaciones: [],
      version: 1,
      deletedAt: null,
      createdAt: ahora,
      updatedAt: ahora,
      historial: [
        {
          fechaHora: formatNowDateTime(),
          usuario: session.nombre,
          rol: rolLabel,
          accion: `Creó el pedido #${numeroPedido}`
        }
      ]
    });

    store.pedidos.unshift(nuevo);
    writeLocalStore(store);
    return nuevo;
  }

  const db = getDb();
  return db.transaction(async (tx) => {
    const seqRes = await tx.execute(sql`SELECT nextval('public.pedidos_numero_seq')::bigint AS num`);
    const rawNum = Number((seqRes[0] as { num?: string | number } | undefined)?.num ?? 1);
    const numeroPedido = String(rawNum).padStart(4, '0');

    const nuevo = recalcularPedido({
      id: randomUUID(),
      numeroPedido,
      cliente: input.cliente.trim(),
      vendedor: input.vendedor,
      whatsapp: input.whatsapp.trim(),
      fechaPedido: input.fechaPedido || hoy,
      despachoPrevisto: input.despachoPrevisto || addDaysISO(hoy, 5),
      lineas: lineasSeguras,
      descuentoPorc: esAdmin ? input.descuentoPorc : 0,
      medioPago: esAdmin ? input.medioPago : 'Sin pagar',
      montoAbonado: esAdmin ? input.montoAbonado : 0,
      montoCheque: esAdmin ? input.montoCheque : 0,
      chequeAcreditado: false,
      fechaRecepcionCheque: esAdmin && input.medioPago === 'Cheque' ? input.fechaRecepcionCheque : '',
      fechaAcreditacionCheque: esAdmin && input.medioPago === 'Cheque' ? input.fechaAcreditacionCheque : '',
      detalleCheque: esAdmin && input.medioPago === 'Cheque' ? input.detalleCheque.trim() : '',
      estadoArmado: 'Por armar',
      estadoEnvio: 'Pendiente',
      transporte: '',
      nroSeguimiento: '',
      observaciones: input.observaciones.trim(),
      createdBy: session.userId,
      incidenciaDeposito: '',
      incidenciaAt: null,
      incidenciaPor: '',
      asignaciones: [],
      version: 1,
      deletedAt: null,
      createdAt: ahora,
      updatedAt: ahora,
      historial: [
        {
          fechaHora: formatNowDateTime(),
          usuario: session.nombre,
          rol: rolLabel,
          accion: `Creó el pedido #${numeroPedido}`
        }
      ]
    });

    await tx.insert(pedidos).values({
      id: nuevo.id,
      numeroPedido: nuevo.numeroPedido,
      cliente: nuevo.cliente,
      vendedor: nuevo.vendedor,
      whatsapp: nuevo.whatsapp,
      fechaPedido: nuevo.fechaPedido,
      despachoPrevisto: nuevo.despachoPrevisto || null,
      descuentoPorc: toNumericString(nuevo.descuentoPorc),
      subtotalNeto: toNumericString(nuevo.subtotalNeto),
      montoDescuento: toNumericString(nuevo.montoDescuento),
      baseImponible: toNumericString(nuevo.baseImponible),
      iva21: toNumericString(nuevo.iva21),
      totalFacturado: toNumericString(nuevo.totalFacturado),
      totalFacturadoOverride: null,
      medioPago: nuevo.medioPago,
      montoAbonado: toNumericString(nuevo.montoAbonado),
      montoCheque: toNumericString(nuevo.montoCheque),
      chequeAcreditado: nuevo.chequeAcreditado,
      fechaRecepcionCheque: nuevo.fechaRecepcionCheque || null,
      fechaAcreditacionCheque: nuevo.fechaAcreditacionCheque || null,
      detalleCheque: nuevo.detalleCheque,
      saldoPendiente: toNumericString(nuevo.saldoPendiente),
      deudaSinRespaldo: toNumericString(nuevo.deudaSinRespaldo),
      estadoPago: nuevo.estadoPago,
      estadoArmado: nuevo.estadoArmado,
      estadoEnvio: nuevo.estadoEnvio,
      transporte: nuevo.transporte,
      nroSeguimiento: nuevo.nroSeguimiento,
      observaciones: nuevo.observaciones,
      createdBy: session.userId,
      incidenciaDeposito: '',
      incidenciaPor: '',
      version: 1
    });

    await tx.insert(pedidoLineas).values(
      nuevo.lineas.map((l) => ({
        pedidoId: nuevo.id,
        idLinea: l.idLinea,
        idGrupo: l.idGrupo,
        nombreProducto: l.nombreProducto,
        rubro: l.rubro,
        precioUnitarioNeto: toNumericString(l.precioUnitarioNeto),
        cantidadTotal: l.cantidadTotal,
        notaCliente: l.notaCliente,
        estado: l.estado,
        requiereRevision: l.requiereRevision,
        motivoRevision: l.motivoRevision ?? null,
        desgloseJson: l.desglose
      }))
    );

    await tx.insert(pedidoHistorial).values({
      pedidoId: nuevo.id,
      usuarioId: session.userId,
      usuarioNombre: session.nombre,
      rol: rolLabel,
      accion: `Creó el pedido #${numeroPedido}`
    });

    return nuevo;
  });
};

const applyPatchToPedido = (
  actual: Pedido,
  patch: PatchPedidoInput,
  actorNombre: string,
  actorRol: ReturnType<typeof rolUsuarioToLabel>,
  motivo: string,
  customCatalogo: ReadonlyArray<ProductoCustomRecord>
): Pedido => {
  const hoy = getTodayISO();
  const nextLineas = patch.lineas
    ? sanitizarLineasContraCatalogo(patch.lineas, customCatalogo)
    : actual.lineas;

  const nextEstadoEnvio = patch.estadoEnvio ?? actual.estadoEnvio;
  let nextFechaDespacho = patch.fechaDespacho !== undefined ? patch.fechaDespacho : actual.fechaDespacho;
  let nextFechaEntrega = patch.fechaEntrega !== undefined ? patch.fechaEntrega : actual.fechaEntrega;

  if (patch.estadoEnvio === 'Despachado' && !nextFechaDespacho) {
    nextFechaDespacho = hoy;
  }
  if (patch.estadoEnvio === 'Entregado') {
    if (!nextFechaDespacho) nextFechaDespacho = hoy;
    if (!nextFechaEntrega) nextFechaEntrega = hoy;
  }

  const evento: HistorialEvento = {
    fechaHora: formatNowDateTime(),
    usuario: actorNombre,
    rol: actorRol,
    accion: motivo
  };

  const merged: PedidoBase = {
    ...actual,
    cliente: patch.cliente !== undefined ? patch.cliente.trim() : actual.cliente,
    vendedor: patch.vendedor !== undefined ? patch.vendedor : actual.vendedor,
    whatsapp: patch.whatsapp !== undefined ? patch.whatsapp.trim() : actual.whatsapp,
    despachoPrevisto: patch.despachoPrevisto !== undefined ? patch.despachoPrevisto : actual.despachoPrevisto,
    descuentoPorc: patch.descuentoPorc !== undefined ? patch.descuentoPorc : actual.descuentoPorc,
    medioPago: patch.medioPago !== undefined ? patch.medioPago : actual.medioPago,
    montoAbonado: patch.montoAbonado !== undefined ? patch.montoAbonado : actual.montoAbonado,
    montoCheque: patch.montoCheque !== undefined ? patch.montoCheque : actual.montoCheque,
    chequeAcreditado: patch.chequeAcreditado !== undefined ? patch.chequeAcreditado : actual.chequeAcreditado,
    fechaRecepcionCheque:
      patch.fechaRecepcionCheque !== undefined ? patch.fechaRecepcionCheque : actual.fechaRecepcionCheque,
    fechaAcreditacionCheque:
      patch.fechaAcreditacionCheque !== undefined ? patch.fechaAcreditacionCheque : actual.fechaAcreditacionCheque,
    detalleCheque: patch.detalleCheque !== undefined ? patch.detalleCheque : actual.detalleCheque,
    estadoArmado: patch.estadoArmado !== undefined ? patch.estadoArmado : actual.estadoArmado,
    estadoEnvio: nextEstadoEnvio,
    fechaDespacho: nextFechaDespacho,
    fechaEntrega: nextFechaEntrega,
    transporte: patch.transporte !== undefined ? patch.transporte : actual.transporte,
    nroSeguimiento: patch.nroSeguimiento !== undefined ? patch.nroSeguimiento : actual.nroSeguimiento,
    observaciones: patch.observaciones !== undefined ? patch.observaciones : actual.observaciones,
    lineas: nextLineas,
    version: actual.version + 1,
    updatedAt: Date.now(),
    historial: [evento, ...actual.historial]
  };

  return recalcularPedido(merged);
};

export const patchPedidoDal = async (
  pedidoId: string,
  patch: PatchPedidoInput,
  motivo: string
): Promise<Pedido> => {
  const session = await requireSession();
  const esAdmin = session.rol === 'ADMIN';
  const esDeposito = isDepositoRole(session.rol);

  if (esDeposito) {
    for (const key of Object.keys(patch) as Array<keyof PatchPedidoInput>) {
      if (patch[key] !== undefined && !DEPOSITO_ALLOWED_PATCH_KEYS.has(key)) {
        throw new AuthError(`Campo "${String(key)}" no permitido para cuenta de Depósito.`, 403);
      }
    }
  } else if (!esAdmin) {
    for (const key of ADMIN_ONLY_PATCH_KEYS) {
      if (patch[key] !== undefined) {
        throw new AuthError(`Campo "${String(key)}" restringido a Administración.`, 403);
      }
    }
  }

  const customCatalogo = patch.lineas ? await listarProductosCustomDal(false) : [];
  const rolLabel = rolUsuarioToLabel(session.rol);

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const idx = store.pedidos.findIndex((p) => p.id === pedidoId && !p.deletedAt);
    if (idx === -1) throw new Error('Pedido no encontrado.');
    const actualWithAsig = attachLocalAsignaciones([store.pedidos[idx]!], store.pedidoAsignaciones)[0]!;
    if (!isPedidoVisibleForSession(actualWithAsig, session)) {
      throw new AuthError('No tenés acceso a este pedido.', 403);
    }

    let actualizado = applyPatchToPedido(
      actualWithAsig,
      patch,
      session.nombre,
      rolLabel,
      motivo,
      customCatalogo
    );

    // Detección inteligente: si ADMIN modifica artículos de un pedido ya enviado a DEPOSITO -> requiere_relectura
    if (esAdmin && patch.lineas !== undefined) {
      let huboCambioEnDeposito = false;
      const now = Date.now();
      for (const asig of store.pedidoAsignaciones) {
        if (asig.pedidoId === pedidoId && isDepositoRole(asig.usuarioRol)) {
          asig.requiereRelectura = true;
          asig.vistoAt = null;
          asig.updatedAt = now;
          huboCambioEnDeposito = true;
        }
      }
      if (huboCambioEnDeposito) {
        actualizado = {
          ...actualizado,
          historial: [
            {
              fechaHora: formatNowDateTime(),
              usuario: session.nombre,
              rol: rolLabel,
              accion: 'Artículos modificados después de enviar a Depósito — requiere revisar hoja'
            },
            ...actualizado.historial
          ]
        };
      }
    }

    const finalAsig = store.pedidoAsignaciones.filter((a) => a.pedidoId === pedidoId);
    store.pedidos[idx] = { ...actualizado, asignaciones: finalAsig };
    writeLocalStore(store);
    return sanitizePedidoForRole({ ...actualizado, asignaciones: finalAsig }, session.rol, session.userId);
  }

  const db = getDb();
  const res = await db.transaction(async (tx) => {
    const rows = await tx
      .select()
      .from(pedidos)
      .where(and(eq(pedidos.id, pedidoId), isNull(pedidos.deletedAt)))
      .limit(1);
    const pr = rows[0];
    if (!pr) throw new Error('Pedido no encontrado.');

    const [lineasRows, historialRows, asignacionesRows] = await Promise.all([
      tx.select().from(pedidoLineas).where(eq(pedidoLineas.pedidoId, pedidoId)),
      tx
        .select()
        .from(pedidoHistorial)
        .where(eq(pedidoHistorial.pedidoId, pedidoId))
        .orderBy(desc(pedidoHistorial.createdAt)),
      tx
        .select({
          asig: pedidoAsignaciones,
          usuarioNombre: usuarios.nombre,
          usuarioRol: usuarios.rol
        })
        .from(pedidoAsignaciones)
        .leftJoin(usuarios, eq(pedidoAsignaciones.usuarioId, usuarios.id))
        .where(eq(pedidoAsignaciones.pedidoId, pedidoId))
    ]);

    const actual = mapDbRowsToPedidos([pr], lineasRows, historialRows, asignacionesRows)[0]!;
    if (!isPedidoVisibleForSession(actual, session)) {
      throw new AuthError('No tenés acceso a este pedido.', 403);
    }

    let actualizado = applyPatchToPedido(
      actual,
      patch,
      session.nombre,
      rolLabel,
      motivo,
      customCatalogo
    );

    await tx
      .update(pedidos)
      .set({
        ...(patch.cliente !== undefined ? { cliente: actualizado.cliente } : {}),
        ...(patch.vendedor !== undefined ? { vendedor: actualizado.vendedor } : {}),
        ...(patch.whatsapp !== undefined ? { whatsapp: actualizado.whatsapp } : {}),
        ...(patch.despachoPrevisto !== undefined
          ? { despachoPrevisto: actualizado.despachoPrevisto || null }
          : {}),
        ...(patch.descuentoPorc !== undefined
          ? { descuentoPorc: toNumericString(actualizado.descuentoPorc) }
          : {}),
        ...(patch.medioPago !== undefined ? { medioPago: actualizado.medioPago } : {}),
        ...(patch.montoAbonado !== undefined || patch.descuentoPorc !== undefined || patch.lineas !== undefined
          ? { montoAbonado: toNumericString(actualizado.montoAbonado) }
          : {}),
        ...(patch.montoCheque !== undefined ||
        patch.medioPago !== undefined ||
        patch.chequeAcreditado !== undefined ||
        patch.montoAbonado !== undefined ||
        patch.descuentoPorc !== undefined ||
        patch.lineas !== undefined
          ? { montoCheque: toNumericString(actualizado.montoCheque) }
          : {}),
        ...(patch.chequeAcreditado !== undefined ? { chequeAcreditado: actualizado.chequeAcreditado } : {}),
        ...(patch.fechaRecepcionCheque !== undefined
          ? { fechaRecepcionCheque: actualizado.fechaRecepcionCheque || null }
          : {}),
        ...(patch.fechaAcreditacionCheque !== undefined
          ? { fechaAcreditacionCheque: actualizado.fechaAcreditacionCheque || null }
          : {}),
        ...(patch.detalleCheque !== undefined ? { detalleCheque: actualizado.detalleCheque } : {}),
        ...(patch.transporte !== undefined ? { transporte: actualizado.transporte } : {}),
        ...(patch.nroSeguimiento !== undefined ? { nroSeguimiento: actualizado.nroSeguimiento } : {}),
        ...(patch.observaciones !== undefined ? { observaciones: actualizado.observaciones } : {}),
        ...(patch.estadoEnvio !== undefined
          ? {
              estadoEnvio: actualizado.estadoEnvio,
              fechaDespacho: actualizado.fechaDespacho || null,
              fechaEntrega: actualizado.fechaEntrega || null
            }
          : {}),
        estadoArmado: actualizado.estadoArmado,
        subtotalNeto: toNumericString(actualizado.subtotalNeto),
        montoDescuento: toNumericString(actualizado.montoDescuento),
        baseImponible: toNumericString(actualizado.baseImponible),
        iva21: toNumericString(actualizado.iva21),
        totalFacturado: toNumericString(actualizado.totalFacturado),
        saldoPendiente: toNumericString(actualizado.saldoPendiente),
        deudaSinRespaldo: toNumericString(actualizado.deudaSinRespaldo),
        estadoPago: actualizado.estadoPago,
        version: actualizado.version,
        updatedAt: new Date(actualizado.updatedAt)
      })
      .where(eq(pedidos.id, pedidoId));

    if (patch.lineas !== undefined) {
      await tx.delete(pedidoLineas).where(eq(pedidoLineas.pedidoId, pedidoId));
      if (actualizado.lineas.length > 0) {
        await tx.insert(pedidoLineas).values(
          actualizado.lineas.map((l) => ({
            pedidoId,
            idLinea: l.idLinea,
            idGrupo: l.idGrupo,
            nombreProducto: l.nombreProducto,
            rubro: l.rubro,
            precioUnitarioNeto: toNumericString(l.precioUnitarioNeto),
            cantidadTotal: l.cantidadTotal,
            notaCliente: l.notaCliente,
            estado: l.estado,
            requiereRevision: l.requiereRevision,
            motivoRevision: l.motivoRevision ?? null,
            desgloseJson: l.desglose
          }))
        );
      }

      if (esAdmin) {
        const depositoAsigIds = actual.asignaciones
          .filter((a) => isDepositoRole(a.usuarioRol))
          .map((a) => a.id);

        if (depositoAsigIds.length > 0) {
          const now = new Date();
          await tx
            .update(pedidoAsignaciones)
            .set({
              requiereRelectura: true,
              vistoAt: null,
              updatedAt: now
            })
            .where(inArray(pedidoAsignaciones.id, depositoAsigIds));

          const notaRelectura = 'Artículos modificados después de enviar a Depósito — requiere revisar hoja';
          await tx.insert(pedidoHistorial).values({
            pedidoId,
            usuarioId: session.userId,
            usuarioNombre: session.nombre,
            rol: rolLabel,
            accion: notaRelectura
          });

          actualizado = {
            ...actualizado,
            asignaciones: actual.asignaciones.map((a) =>
              depositoAsigIds.includes(a.id)
                ? { ...a, requiereRelectura: true, vistoAt: null, updatedAt: now.getTime() }
                : a
            ),
            historial: [
              {
                fechaHora: formatNowDateTime(),
                usuario: session.nombre,
                rol: rolLabel,
                accion: notaRelectura
              },
              ...actualizado.historial
            ]
          };
        }
      }
    }

    await tx.insert(pedidoHistorial).values({
      pedidoId,
      usuarioId: session.userId,
      usuarioNombre: session.nombre,
      rol: rolLabel,
      accion: motivo
    });

    return actualizado;
  });

  return sanitizePedidoForRole(res, session.rol, session.userId);
};

export const enviarPedidosACuentasDal = async (
  input: EnviarCuentasInput
): Promise<EnviarCuentasResult> => {
  const session = await requireAdmin();
  const rolLabel = rolUsuarioToLabel(session.rol);
  const pedidoIds = Array.from(new Set(input.pedidoIds));
  const usuarioIds = Array.from(new Set(input.usuarioIds));

  if (pedidoIds.length === 0 || usuarioIds.length === 0) {
    return {
      ok: false,
      enviados: 0,
      omitidosTalle: [],
      omitidosSinLineas: [],
      error: 'Seleccioná al menos un pedido y una cuenta de destino.'
    };
  }

  const notaEnvio = input.notaEnvio.trim();
  const prioridad: PrioridadAsignacion = input.prioridad === 'URGENTE' ? 'URGENTE' : 'NORMAL';

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const cuentas = store.usuarios.filter(
      (u) => usuarioIds.includes(u.id) && u.activo && u.rol !== 'ADMIN'
    );
    if (cuentas.length === 0) {
      return {
        ok: false,
        enviados: 0,
        omitidosTalle: [],
        omitidosSinLineas: [],
        error: 'Las cuentas seleccionadas no existen o están inactivas.'
      };
    }

    const incluyeDeposito = cuentas.some((c) => isDepositoRole(c.rol));
    const targetPedidos = store.pedidos.filter((p) => !p.deletedAt && pedidoIds.includes(p.id));
    const omitidosTalle: string[] = [];
    const omitidosSinLineas: string[] = [];
    const validos: Pedido[] = [];

    for (const p of targetPedidos) {
      if (incluyeDeposito) {
        if (p.faltantesVariante > 0) {
          omitidosTalle.push(`#${p.numeroPedido}`);
          continue;
        }
        if (p.lineas.length === 0) {
          omitidosSinLineas.push(`#${p.numeroPedido}`);
          continue;
        }
      }
      validos.push(p);
    }

    if (validos.length === 0) {
      const msg =
        omitidosTalle.length > 0
          ? 'No se puede enviar a Depósito porque falta definir talle o color. Resolvelo antes o derivalo a Ventas.'
          : 'No se puede enviar a Depósito un pedido sin artículos desglosados para armar.';
      return {
        ok: false,
        enviados: 0,
        omitidosTalle,
        omitidosSinLineas,
        error: msg
      };
    }

    if (incluyeDeposito) {
      const sinPago = validos
        .filter((p) => p.estadoPago === 'Falta pago')
        .map((p) => `#${p.numeroPedido}`);
      if (sinPago.length > 0 && !input.forzarSinPago) {
        return {
          ok: false,
          enviados: 0,
          omitidosTalle,
          omitidosSinLineas,
          requiereConfirmacionSinPago: true,
          pedidosSinPago: sinPago
        };
      }
    }

    const now = Date.now();
    const nombresCuentas = cuentas.map((c) => c.nombre).join(', ');

    for (const p of validos) {
      for (const cta of cuentas) {
        const existingIdx = store.pedidoAsignaciones.findIndex(
          (a) => a.pedidoId === p.id && a.usuarioId === cta.id
        );
        if (existingIdx !== -1) {
          const prev = store.pedidoAsignaciones[existingIdx]!;
          store.pedidoAsignaciones[existingIdx] = {
            ...prev,
            usuarioNombre: cta.nombre,
            usuarioRol: cta.rol,
            enviadoPor: session.userId,
            enviadoPorNombre: session.nombre,
            notaEnvio,
            prioridad,
            vistoAt: null,
            requiereRelectura: false,
            updatedAt: now
          };
        } else {
          store.pedidoAsignaciones.push({
            id: randomUUID(),
            pedidoId: p.id,
            usuarioId: cta.id,
            usuarioNombre: cta.nombre,
            usuarioRol: cta.rol,
            enviadoPor: session.userId,
            enviadoPorNombre: session.nombre,
            notaEnvio,
            prioridad,
            vistoAt: null,
            requiereRelectura: false,
            createdAt: now,
            updatedAt: now
          });
        }
      }

      const pIdx = store.pedidos.findIndex((x) => x.id === p.id);
      if (pIdx !== -1) {
        const cur = store.pedidos[pIdx]!;
        cur.updatedAt = now;
        cur.historial.unshift({
          fechaHora: formatNowDateTime(),
          usuario: session.nombre,
          rol: rolLabel,
          accion: `Envió pedido a ${nombresCuentas} (${prioridad === 'URGENTE' ? 'Urgente' : 'Normal'})${
            notaEnvio ? ` · Nota: ${notaEnvio}` : ''
          }`
        });
      }
    }

    writeLocalStore(store);
    return {
      ok: true,
      enviados: validos.length,
      omitidosTalle,
      omitidosSinLineas
    };
  }

  const db = getDb();
  return db.transaction(async (tx) => {
    const cuentas = await tx
      .select()
      .from(usuarios)
      .where(and(inArray(usuarios.id, usuarioIds), eq(usuarios.activo, true)));
    const cuentasValidas = cuentas.filter((c) => c.rol !== 'ADMIN');
    if (cuentasValidas.length === 0) {
      return {
        ok: false,
        enviados: 0,
        omitidosTalle: [],
        omitidosSinLineas: [],
        error: 'Las cuentas seleccionadas no existen o están inactivas.'
      };
    }

    const incluyeDeposito = cuentasValidas.some((c) => isDepositoRole(c.rol));
    const pedidoRows = await tx
      .select()
      .from(pedidos)
      .where(and(inArray(pedidos.id, pedidoIds), isNull(pedidos.deletedAt)));

    const foundIds = pedidoRows.map((p) => p.id);
    const lineasRows =
      foundIds.length > 0
        ? await tx.select().from(pedidoLineas).where(inArray(pedidoLineas.pedidoId, foundIds))
        : [];
    const targetPedidos = mapDbRowsToPedidos(pedidoRows, lineasRows, [], []);

    const omitidosTalle: string[] = [];
    const omitidosSinLineas: string[] = [];
    const validos: Pedido[] = [];

    for (const p of targetPedidos) {
      if (incluyeDeposito) {
        if (p.faltantesVariante > 0) {
          omitidosTalle.push(`#${p.numeroPedido}`);
          continue;
        }
        if (p.lineas.length === 0) {
          omitidosSinLineas.push(`#${p.numeroPedido}`);
          continue;
        }
      }
      validos.push(p);
    }

    if (validos.length === 0) {
      const msg =
        omitidosTalle.length > 0
          ? 'No se puede enviar a Depósito porque falta definir talle o color. Resolvelo antes o derivalo a Ventas.'
          : 'No se puede enviar a Depósito un pedido sin artículos desglosados para armar.';
      return {
        ok: false,
        enviados: 0,
        omitidosTalle,
        omitidosSinLineas,
        error: msg
      };
    }

    if (incluyeDeposito) {
      const sinPago = validos
        .filter((p) => p.estadoPago === 'Falta pago')
        .map((p) => `#${p.numeroPedido}`);
      if (sinPago.length > 0 && !input.forzarSinPago) {
        return {
          ok: false,
          enviados: 0,
          omitidosTalle,
          omitidosSinLineas,
          requiereConfirmacionSinPago: true,
          pedidosSinPago: sinPago
        };
      }
    }

    const now = new Date();
    const nombresCuentas = cuentasValidas.map((c) => c.nombre).join(', ');

    for (const p of validos) {
      for (const cta of cuentasValidas) {
        await tx
          .insert(pedidoAsignaciones)
          .values({
            pedidoId: p.id,
            usuarioId: cta.id,
            enviadoPor: session.userId,
            enviadoPorNombre: session.nombre,
            notaEnvio,
            prioridad,
            vistoAt: null,
            requiereRelectura: false,
            updatedAt: now
          })
          .onConflictDoUpdate({
            target: [pedidoAsignaciones.pedidoId, pedidoAsignaciones.usuarioId],
            set: {
              enviadoPor: session.userId,
              enviadoPorNombre: session.nombre,
              notaEnvio,
              prioridad,
              vistoAt: null,
              requiereRelectura: false,
              updatedAt: now
            }
          });
      }

      await tx.update(pedidos).set({ updatedAt: now }).where(eq(pedidos.id, p.id));
      await tx.insert(pedidoHistorial).values({
        pedidoId: p.id,
        usuarioId: session.userId,
        usuarioNombre: session.nombre,
        rol: rolLabel,
        accion: `Envió pedido a ${nombresCuentas} (${prioridad === 'URGENTE' ? 'Urgente' : 'Normal'})${
          notaEnvio ? ` · Nota: ${notaEnvio}` : ''
        }`
      });
    }

    return {
      ok: true,
      enviados: validos.length,
      omitidosTalle,
      omitidosSinLineas
    };
  });
};

export const retirarAsignacionDal = async (pedidoId: string, usuarioId: string): Promise<void> => {
  const session = await requireAdmin();
  const rolLabel = rolUsuarioToLabel(session.rol);

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const idx = store.pedidoAsignaciones.findIndex(
      (a) => a.pedidoId === pedidoId && a.usuarioId === usuarioId
    );
    if (idx === -1) return;
    const removed = store.pedidoAsignaciones[idx]!;
    store.pedidoAsignaciones.splice(idx, 1);

    const p = store.pedidos.find((x) => x.id === pedidoId);
    if (p) {
      p.updatedAt = Date.now();
      p.historial.unshift({
        fechaHora: formatNowDateTime(),
        usuario: session.nombre,
        rol: rolLabel,
        accion: `Retiró el envío asignado a ${removed.usuarioNombre}`
      });
    }
    writeLocalStore(store);
    return;
  }

  const db = getDb();
  await db.transaction(async (tx) => {
    const targetUser = await tx
      .select({ nombre: usuarios.nombre })
      .from(usuarios)
      .where(eq(usuarios.id, usuarioId))
      .limit(1);

    await tx
      .delete(pedidoAsignaciones)
      .where(and(eq(pedidoAsignaciones.pedidoId, pedidoId), eq(pedidoAsignaciones.usuarioId, usuarioId)));

    await tx.update(pedidos).set({ updatedAt: new Date() }).where(eq(pedidos.id, pedidoId));
    await tx.insert(pedidoHistorial).values({
      pedidoId,
      usuarioId: session.userId,
      usuarioNombre: session.nombre,
      rol: rolLabel,
      accion: `Retiró el envío asignado a ${targetUser[0]?.nombre ?? 'cuenta'}`
    });
  });
};

export const marcarPedidoVistoDal = async (pedidoId: string): Promise<void> => {
  const session = await requireSession();
  if (session.rol === 'ADMIN') return;

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const asig = store.pedidoAsignaciones.find(
      (a) => a.pedidoId === pedidoId && a.usuarioId === session.userId
    );
    if (!asig) return;
    if (asig.vistoAt !== null && !asig.requiereRelectura) return;
    const now = Date.now();
    asig.vistoAt = now;
    asig.requiereRelectura = false;
    asig.updatedAt = now;
    const p = store.pedidos.find((x) => x.id === pedidoId);
    if (p) p.updatedAt = now;
    writeLocalStore(store);
    return;
  }

  const db = getDb();
  const now = new Date();
  await db
    .update(pedidoAsignaciones)
    .set({
      vistoAt: now,
      requiereRelectura: false,
      updatedAt: now
    })
    .where(
      and(eq(pedidoAsignaciones.pedidoId, pedidoId), eq(pedidoAsignaciones.usuarioId, session.userId))
    );
  await db.update(pedidos).set({ updatedAt: now }).where(eq(pedidos.id, pedidoId));
};

export const reportarIncidenciaDepositoDal = async (
  pedidoId: string,
  mensaje: string
): Promise<Pedido> => {
  const session = await requireSession();
  if (!isDepositoRole(session.rol) && session.rol !== 'ADMIN') {
    throw new AuthError('Solo Depósito o Administración pueden reportar incidencias de armado.', 403);
  }
  const texto = mensaje.trim();
  if (!texto) throw new Error('Describí el faltante o problema antes de enviar el aviso.');
  const rolLabel = rolUsuarioToLabel(session.rol);

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const idx = store.pedidos.findIndex((p) => p.id === pedidoId && !p.deletedAt);
    if (idx === -1) throw new Error('Pedido no encontrado.');
    const withAsig = attachLocalAsignaciones([store.pedidos[idx]!], store.pedidoAsignaciones)[0]!;
    if (!isPedidoVisibleForSession(withAsig, session)) {
      throw new AuthError('No tenés acceso a este pedido.', 403);
    }
    const now = Date.now();
    const updated = recalcularPedido({
      ...withAsig,
      incidenciaDeposito: texto,
      incidenciaAt: now,
      incidenciaPor: session.nombre,
      updatedAt: now,
      version: withAsig.version + 1,
      historial: [
        {
          fechaHora: formatNowDateTime(),
          usuario: session.nombre,
          rol: rolLabel,
          accion: `Reportó problema en armado: ${texto}`
        },
        ...withAsig.historial
      ]
    });
    store.pedidos[idx] = updated;
    writeLocalStore(store);
    return sanitizePedidoForRole(updated, session.rol, session.userId);
  }

  const db = getDb();
  const updated = await db.transaction(async (tx) => {
    const existing = await tx
      .select()
      .from(pedidos)
      .where(and(eq(pedidos.id, pedidoId), isNull(pedidos.deletedAt)))
      .limit(1);
    const pr = existing[0];
    if (!pr) throw new Error('Pedido no encontrado.');

    if (isDepositoRole(session.rol)) {
      const asig = await tx
        .select({ id: pedidoAsignaciones.id })
        .from(pedidoAsignaciones)
        .where(
          and(eq(pedidoAsignaciones.pedidoId, pedidoId), eq(pedidoAsignaciones.usuarioId, session.userId))
        )
        .limit(1);
      if (asig.length === 0) {
        throw new AuthError('No tenés acceso a este pedido.', 403);
      }
    }

    const now = new Date();
    await tx
      .update(pedidos)
      .set({
        incidenciaDeposito: texto,
        incidenciaAt: now,
        incidenciaPor: session.nombre,
        version: pr.version + 1,
        updatedAt: now
      })
      .where(eq(pedidos.id, pedidoId));

    await tx.insert(pedidoHistorial).values({
      pedidoId,
      usuarioId: session.userId,
      usuarioNombre: session.nombre,
      rol: rolLabel,
      accion: `Reportó problema en armado: ${texto}`
    });

    const [pedidoRows, lineasRows, historialRows, asignacionesRows] = await Promise.all([
      tx.select().from(pedidos).where(eq(pedidos.id, pedidoId)).limit(1),
      tx.select().from(pedidoLineas).where(eq(pedidoLineas.pedidoId, pedidoId)),
      tx
        .select()
        .from(pedidoHistorial)
        .where(eq(pedidoHistorial.pedidoId, pedidoId))
        .orderBy(desc(pedidoHistorial.createdAt)),
      tx
        .select({
          asig: pedidoAsignaciones,
          usuarioNombre: usuarios.nombre,
          usuarioRol: usuarios.rol
        })
        .from(pedidoAsignaciones)
        .leftJoin(usuarios, eq(pedidoAsignaciones.usuarioId, usuarios.id))
        .where(eq(pedidoAsignaciones.pedidoId, pedidoId))
    ]);

    return mapDbRowsToPedidos(pedidoRows, lineasRows, historialRows, asignacionesRows)[0]!;
  });

  return sanitizePedidoForRole(updated, session.rol, session.userId);
};

export const resolverIncidenciaDepositoDal = async (pedidoId: string): Promise<Pedido> => {
  const session = await requireAdmin();
  const rolLabel = rolUsuarioToLabel(session.rol);

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const idx = store.pedidos.findIndex((p) => p.id === pedidoId && !p.deletedAt);
    if (idx === -1) throw new Error('Pedido no encontrado.');
    const withAsig = attachLocalAsignaciones([store.pedidos[idx]!], store.pedidoAsignaciones)[0]!;
    const now = Date.now();
    const updated = recalcularPedido({
      ...withAsig,
      incidenciaDeposito: '',
      incidenciaAt: null,
      incidenciaPor: '',
      updatedAt: now,
      version: withAsig.version + 1,
      historial: [
        {
          fechaHora: formatNowDateTime(),
          usuario: session.nombre,
          rol: rolLabel,
          accion: 'Marcó como resuelto el aviso de Depósito'
        },
        ...withAsig.historial
      ]
    });
    store.pedidos[idx] = updated;
    writeLocalStore(store);
    return updated;
  }

  const db = getDb();
  return db.transaction(async (tx) => {
    const existing = await tx
      .select()
      .from(pedidos)
      .where(and(eq(pedidos.id, pedidoId), isNull(pedidos.deletedAt)))
      .limit(1);
    const pr = existing[0];
    if (!pr) throw new Error('Pedido no encontrado.');

    const now = new Date();
    await tx
      .update(pedidos)
      .set({
        incidenciaDeposito: '',
        incidenciaAt: null,
        incidenciaPor: '',
        version: pr.version + 1,
        updatedAt: now
      })
      .where(eq(pedidos.id, pedidoId));

    await tx.insert(pedidoHistorial).values({
      pedidoId,
      usuarioId: session.userId,
      usuarioNombre: session.nombre,
      rol: rolLabel,
      accion: 'Marcó como resuelto el aviso de Depósito'
    });

    const [pedidoRows, lineasRows, historialRows, asignacionesRows] = await Promise.all([
      tx.select().from(pedidos).where(eq(pedidos.id, pedidoId)).limit(1),
      tx.select().from(pedidoLineas).where(eq(pedidoLineas.pedidoId, pedidoId)),
      tx
        .select()
        .from(pedidoHistorial)
        .where(eq(pedidoHistorial.pedidoId, pedidoId))
        .orderBy(desc(pedidoHistorial.createdAt)),
      tx
        .select({
          asig: pedidoAsignaciones,
          usuarioNombre: usuarios.nombre,
          usuarioRol: usuarios.rol
        })
        .from(pedidoAsignaciones)
        .leftJoin(usuarios, eq(pedidoAsignaciones.usuarioId, usuarios.id))
        .where(eq(pedidoAsignaciones.pedidoId, pedidoId))
    ]);

    return mapDbRowsToPedidos(pedidoRows, lineasRows, historialRows, asignacionesRows)[0]!;
  });
};

export const batchUpdatePedidosDal = async (
  ids: ReadonlyArray<string>,
  action: BatchActionInput
): Promise<{ actualizados: number; omitidos: string[] }> => {
  const session =
    action.type === 'pago' || action.type === 'vendedor' ? await requireAdmin() : await requireSession();
  const rolLabel = rolUsuarioToLabel(session.rol);
  const hoy = getTodayISO();

  const omitidos: string[] = [];
  let actualizados = 0;

  const targetPedidos = await obtenerPedidosPorIdsDal(ids);
  for (const p of targetPedidos) {
    if (action.type === 'armado') {
      if (action.estadoArmado === 'Listo' && p.faltantesVariante > 0) {
        omitidos.push(`#${p.numeroPedido}`);
        continue;
      }
      await patchPedidoDal(
        p.id,
        { estadoArmado: action.estadoArmado },
        `Acción en lote (${rolLabel}): pasó armado a ${action.estadoArmado}`
      );
      actualizados += 1;
    } else if (action.type === 'envio') {
      const transp = action.transporte?.trim();
      await patchPedidoDal(
        p.id,
        {
          estadoEnvio: action.estadoEnvio,
          ...(transp ? { transporte: transp } : {})
        },
        transp
          ? `Acción en lote: ${action.estadoEnvio} por ${transp}`
          : `Acción en lote: envío marcado como ${action.estadoEnvio}`
      );
      actualizados += 1;
    } else if (action.type === 'pago') {
      if (action.estado === 'Pagado') {
        const nota =
          p.medioPago === 'Cheque'
            ? `Cheque acreditado en lote el ${formatShortDate(hoy)}`
            : `Cobro confirmado en lote el ${formatShortDate(hoy)}`;
        await patchPedidoDal(
          p.id,
          {
            medioPago: 'Transferencia / Efectivo',
            chequeAcreditado: true,
            montoAbonado: p.totalFacturado,
            montoCheque: 0,
            observaciones: [p.observaciones, nota].filter(Boolean).join(' · ')
          },
          'Acción en lote: marcado como Pagado'
        );
      } else {
        await patchPedidoDal(
          p.id,
          {
            medioPago: 'Sin pagar',
            chequeAcreditado: false,
            montoAbonado: 0,
            montoCheque: 0
          },
          'Acción en lote: pasado a Falta pago'
        );
      }
      actualizados += 1;
    } else if (action.type === 'vendedor') {
      const vend = action.vendedor.trim();
      if (!vend) continue;
      await patchPedidoDal(p.id, { vendedor: vend }, `Acción en lote: vendedor asignado a ${vend}`);
      actualizados += 1;
    }
  }

  return { actualizados, omitidos };
};

export const softDeletePedidosDal = async (ids: ReadonlyArray<string>): Promise<number> => {
  const session = await requireAdmin();
  if (ids.length === 0) return 0;
  const rolLabel = rolUsuarioToLabel(session.rol);
  const nowIso = new Date().toISOString();

  if (isLocalFallbackMode()) {
    const idSet = new Set(ids);
    const store = readLocalStore();
    let count = 0;
    for (const p of store.pedidos) {
      if (idSet.has(p.id) && !p.deletedAt) {
        p.deletedAt = nowIso;
        p.updatedAt = Date.now();
        p.historial.unshift({
          fechaHora: formatNowDateTime(),
          usuario: session.nombre,
          rol: rolLabel,
          accion: 'Eliminó el pedido (ventana Deshacer)'
        });
        count += 1;
      }
    }
    writeLocalStore(store);
    return count;
  }

  const db = getDb();
  return db.transaction(async (tx) => {
    const updated = await tx
      .update(pedidos)
      .set({ deletedAt: new Date(), updatedAt: new Date() })
      .where(and(inArray(pedidos.id, [...ids]), isNull(pedidos.deletedAt)))
      .returning({ id: pedidos.id });

    if (updated.length > 0) {
      await tx.insert(pedidoHistorial).values(
        updated.map((u) => ({
          pedidoId: u.id,
          usuarioId: session.userId,
          usuarioNombre: session.nombre,
          rol: rolLabel,
          accion: 'Eliminó el pedido'
        }))
      );
    }
    return updated.length;
  });
};

export const restorePedidosDal = async (ids: ReadonlyArray<string>): Promise<number> => {
  const session = await requireAdmin();
  if (ids.length === 0) return 0;
  const rolLabel = rolUsuarioToLabel(session.rol);

  if (isLocalFallbackMode()) {
    const idSet = new Set(ids);
    const store = readLocalStore();
    let count = 0;
    for (const p of store.pedidos) {
      if (idSet.has(p.id) && p.deletedAt) {
        p.deletedAt = null;
        p.updatedAt = Date.now();
        p.historial.unshift({
          fechaHora: formatNowDateTime(),
          usuario: session.nombre,
          rol: rolLabel,
          accion: 'Restauró el pedido (Deshacer)'
        });
        count += 1;
      }
    }
    writeLocalStore(store);
    return count;
  }

  const db = getDb();
  return db.transaction(async (tx) => {
    const restored = await tx
      .update(pedidos)
      .set({ deletedAt: null, updatedAt: new Date() })
      .where(inArray(pedidos.id, [...ids]))
      .returning({ id: pedidos.id });

    if (restored.length > 0) {
      await tx.insert(pedidoHistorial).values(
        restored.map((u) => ({
          pedidoId: u.id,
          usuarioId: session.userId,
          usuarioNombre: session.nombre,
          rol: rolLabel,
          accion: 'Restauró el pedido (Deshacer)'
        }))
      );
    }
    return restored.length;
  });
};

export const importarPedidosHistoricosDal = async (
  nuevos: ReadonlyArray<Pedido>
): Promise<{ importados: number; duplicadosOmitidos: number }> => {
  const session = await requireAdmin();
  const rolLabel = rolUsuarioToLabel(session.rol);

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const claves = new Set(
      store.pedidos
        .filter((p) => !p.deletedAt)
        .map((p) => `${p.numeroPedido.replace(/^0+/, '')}|${normalizarTexto(p.cliente)}`)
    );
    const nums = new Set(store.pedidos.map((p) => p.numeroPedido));
    let importados = 0;
    let duplicadosOmitidos = 0;

    for (const p of nuevos) {
      const clave = `${p.numeroPedido.replace(/^0+/, '')}|${normalizarTexto(p.cliente)}`;
      if (claves.has(clave) || nums.has(p.numeroPedido)) {
        duplicadosOmitidos += 1;
        continue;
      }
      claves.add(clave);
      nums.add(p.numeroPedido);
      const id = randomUUID();
      store.pedidos.unshift({
        ...p,
        id,
        createdBy: session.userId,
        incidenciaDeposito: '',
        incidenciaAt: null,
        incidenciaPor: '',
        asignaciones: [],
        historial: [
          {
            fechaHora: formatNowDateTime(),
            usuario: session.nombre,
            rol: rolLabel,
            accion: 'Importado desde planilla general'
          }
        ]
      });
      importados += 1;
    }

    const maxActual = store.pedidos.reduce((acc, p) => {
      const n = parseInt(p.numeroPedido.replace(/\D/g, ''), 10);
      return Number.isFinite(n) && n > acc ? n : acc;
    }, store.nextNumeroSeq);
    store.nextNumeroSeq = maxActual;
    writeLocalStore(store);
    return { importados, duplicadosOmitidos };
  }

  const db = getDb();
  return db.transaction(async (tx) => {
    const existentes = await tx
      .select({ numeroPedido: pedidos.numeroPedido, cliente: pedidos.cliente })
      .from(pedidos);

    const claves = new Set(
      existentes.map((p) => `${p.numeroPedido.replace(/^0+/, '')}|${normalizarTexto(p.cliente)}`)
    );
    const nums = new Set(existentes.map((p) => p.numeroPedido));

    let importados = 0;
    let duplicadosOmitidos = 0;

    for (const p of nuevos) {
      const clave = `${p.numeroPedido.replace(/^0+/, '')}|${normalizarTexto(p.cliente)}`;
      if (claves.has(clave) || nums.has(p.numeroPedido)) {
        duplicadosOmitidos += 1;
        continue;
      }
      claves.add(clave);
      nums.add(p.numeroPedido);

      const newId = randomUUID();
      await tx.insert(pedidos).values({
        id: newId,
        numeroPedido: p.numeroPedido,
        cliente: p.cliente,
        vendedor: p.vendedor,
        whatsapp: p.whatsapp,
        fechaPedido: p.fechaPedido,
        despachoPrevisto: p.despachoPrevisto || null,
        fechaDespacho: p.fechaDespacho || null,
        fechaEntrega: p.fechaEntrega || null,
        descuentoPorc: toNumericString(p.descuentoPorc),
        subtotalNeto: toNumericString(p.subtotalNeto),
        montoDescuento: toNumericString(p.montoDescuento),
        baseImponible: toNumericString(p.baseImponible),
        iva21: toNumericString(p.iva21),
        totalFacturado: toNumericString(p.totalFacturado),
        totalFacturadoOverride:
          p.totalFacturadoOverride !== undefined ? toNumericString(p.totalFacturadoOverride) : null,
        medioPago: p.medioPago,
        montoAbonado: toNumericString(p.montoAbonado),
        montoCheque: toNumericString(p.montoCheque),
        chequeAcreditado: p.chequeAcreditado,
        fechaRecepcionCheque: p.fechaRecepcionCheque || null,
        fechaAcreditacionCheque: p.fechaAcreditacionCheque || null,
        detalleCheque: p.detalleCheque,
        saldoPendiente: toNumericString(p.saldoPendiente),
        deudaSinRespaldo: toNumericString(p.deudaSinRespaldo),
        estadoPago: p.estadoPago,
        estadoArmado: p.estadoArmado,
        estadoEnvio: p.estadoEnvio,
        transporte: p.transporte,
        nroSeguimiento: p.nroSeguimiento,
        observaciones: p.observaciones,
        detalleManualImportado: p.detalleManualImportado ?? null,
        createdBy: session.userId,
        incidenciaDeposito: '',
        incidenciaPor: '',
        version: 1
      });

      await tx.insert(pedidoHistorial).values({
        pedidoId: newId,
        usuarioId: session.userId,
        usuarioNombre: session.nombre,
        rol: rolLabel,
        accion: 'Importado desde planilla general'
      });

      importados += 1;
    }

    await tx.execute(sql`SELECT public.sync_pedidos_numero_seq()`);
    return { importados, duplicadosOmitidos };
  });
};

export const reiniciarSistemaDal = async (): Promise<void> => {
  await requireAdmin();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    store.pedidos = [];
    store.pedidoAsignaciones = [];
    store.nextNumeroSeq = 0;
    writeLocalStore(store);
    return;
  }

  const db = getDb();
  await db.transaction(async (tx) => {
    await tx.delete(pedidos);
    await tx.execute(sql`SELECT setval('public.pedidos_numero_seq', 1, false)`);
  });
};
