import 'server-only';

import { randomUUID } from 'node:crypto';
import { and, desc, eq, inArray, isNull, sql } from 'drizzle-orm';
import { CATALOGO_BY_GRUPO } from '@/lib/catalogo-mibi';
import { getDb, isLocalFallbackMode, readLocalStore, writeLocalStore } from '@/lib/db';
import { pedidoHistorial, pedidoLineas, pedidos } from '@/lib/db/schema';
import {
  addDaysISO,
  formatNowDateTime,
  formatShortDate,
  getTodayISO,
  recalcularPedido
} from '@/lib/finanzas';
import { normalizarTexto } from '@/lib/excel-parser';
import {
  rolUsuarioToLabel,
  type EstadoArmado,
  type EstadoEnvio,
  type HistorialEvento,
  type LineaPedido,
  type MedioPago,
  type Pedido,
  type PedidoBase
} from '@/lib/types';
import { AuthError, requireAdmin, requireSession } from './auth';

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

const toNumericString = (val: number): string => val.toFixed(2);
const fromNumericString = (val: string | null | undefined): number => {
  if (!val) return 0;
  const n = Number(val);
  return Number.isFinite(n) ? n : 0;
};

const sanitizarLineasContraCatalogo = (lineas: ReadonlyArray<LineaPedido>): LineaPedido[] => {
  return lineas.map((l) => {
    const prod = CATALOGO_BY_GRUPO.get(l.idGrupo);
    if (!prod) {
      throw new Error(`Producto desconocido en catálogo: ${l.idGrupo}`);
    }
    const sumaDesglose = l.desglose.reduce((acc, d) => acc + Math.max(0, Math.trunc(d.cantidad)), 0);
    const desgloseValido = l.estado === 'OK' && !l.requiereRevision && sumaDesglose === l.cantidadTotal;
    return {
      ...l,
      nombreProducto: prod.nombre,
      rubro: prod.rubro,
      precioUnitarioNeto: prod.precioNeto,
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

const mapDbRowsToPedidos = (
  pedidoRows: Array<typeof pedidos.$inferSelect>,
  lineasRows: Array<typeof pedidoLineas.$inferSelect>,
  historialRows: Array<typeof pedidoHistorial.$inferSelect>
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
      rol: hr.rol === 'Administración' ? 'Administración' : 'Ventas / Depósito',
      accion: hr.accion
    });
    histByPedido.set(hr.pedidoId, arr);
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

export const listarPedidosDal = async (): Promise<Pedido[]> => {
  await requireSession();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    return sortPedidos(store.pedidos.filter((p) => !p.deletedAt).map((p) => recalcularPedido(p)));
  }

  const db = getDb();
  const pedidoRows = await db
    .select()
    .from(pedidos)
    .where(isNull(pedidos.deletedAt))
    .orderBy(desc(pedidos.createdAt));

  if (pedidoRows.length === 0) return [];
  const ids = pedidoRows.map((p) => p.id);

  const [lineasRows, historialRows] = await Promise.all([
    db.select().from(pedidoLineas).where(inArray(pedidoLineas.pedidoId, ids)),
    db
      .select()
      .from(pedidoHistorial)
      .where(inArray(pedidoHistorial.pedidoId, ids))
      .orderBy(desc(pedidoHistorial.createdAt))
  ]);

  return mapDbRowsToPedidos(pedidoRows, lineasRows, historialRows);
};

export const obtenerPedidoPorIdDal = async (id: string): Promise<Pedido | null> => {
  await requireSession();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const found = store.pedidos.find((p) => p.id === id && !p.deletedAt);
    return found ? recalcularPedido(found) : null;
  }

  const db = getDb();
  const pedidoRows = await db
    .select()
    .from(pedidos)
    .where(and(eq(pedidos.id, id), isNull(pedidos.deletedAt)))
    .limit(1);

  if (pedidoRows.length === 0) return null;
  const [lineasRows, historialRows] = await Promise.all([
    db.select().from(pedidoLineas).where(eq(pedidoLineas.pedidoId, id)),
    db
      .select()
      .from(pedidoHistorial)
      .where(eq(pedidoHistorial.pedidoId, id))
      .orderBy(desc(pedidoHistorial.createdAt))
  ]);

  return mapDbRowsToPedidos(pedidoRows, lineasRows, historialRows)[0] ?? null;
};

export const obtenerPedidosPorIdsDal = async (ids: ReadonlyArray<string>): Promise<Pedido[]> => {
  await requireSession();
  if (ids.length === 0) return [];

  if (isLocalFallbackMode()) {
    const idSet = new Set(ids);
    const store = readLocalStore();
    return sortPedidos(
      store.pedidos.filter((p) => !p.deletedAt && idSet.has(p.id)).map((p) => recalcularPedido(p))
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
  const [lineasRows, historialRows] = await Promise.all([
    db.select().from(pedidoLineas).where(inArray(pedidoLineas.pedidoId, foundIds)),
    db
      .select()
      .from(pedidoHistorial)
      .where(inArray(pedidoHistorial.pedidoId, foundIds))
      .orderBy(desc(pedidoHistorial.createdAt))
  ]);
  return mapDbRowsToPedidos(pedidoRows, lineasRows, historialRows);
};

export const getSyncStatusDal = async (): Promise<{ maxUpdatedAt: number; activeCount: number }> => {
  await requireSession();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const activos = store.pedidos.filter((p) => !p.deletedAt);
    const maxUpdatedAt = activos.reduce((acc, p) => Math.max(acc, p.updatedAt), 0);
    return { maxUpdatedAt, activeCount: activos.length };
  }

  const db = getDb();
  const rows = await db
    .select({
      maxUpdated: sql<Date | null>`max(${pedidos.updatedAt})`,
      count: sql<number>`count(*)::int`
    })
    .from(pedidos)
    .where(isNull(pedidos.deletedAt));

  const row = rows[0];
  const maxUpdatedAt = row?.maxUpdated ? new Date(row.maxUpdated).getTime() : 0;
  return { maxUpdatedAt, activeCount: row?.count ?? 0 };
};

export const crearPedidoDal = async (input: CrearPedidoInput): Promise<Pedido> => {
  const session = await requireSession();
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

  const lineasSeguras = sanitizarLineasContraCatalogo(input.lineas);
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
  actorRol: 'Administración' | 'Ventas / Depósito',
  motivo: string
): Pedido => {
  const hoy = getTodayISO();
  const nextLineas = patch.lineas ? sanitizarLineasContraCatalogo(patch.lineas) : actual.lineas;

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

  if (!esAdmin) {
    for (const key of ADMIN_ONLY_PATCH_KEYS) {
      if (patch[key] !== undefined) {
        throw new AuthError(`Campo "${String(key)}" restringido a Administración.`, 403);
      }
    }
  }

  const rolLabel = rolUsuarioToLabel(session.rol);

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const idx = store.pedidos.findIndex((p) => p.id === pedidoId && !p.deletedAt);
    if (idx === -1) throw new Error('Pedido no encontrado.');
    const actual = recalcularPedido(store.pedidos[idx]!);
    const actualizado = applyPatchToPedido(actual, patch, session.nombre, rolLabel, motivo);
    store.pedidos[idx] = actualizado;
    writeLocalStore(store);
    return actualizado;
  }

  const db = getDb();
  return db.transaction(async (tx) => {
    const rows = await tx
      .select()
      .from(pedidos)
      .where(and(eq(pedidos.id, pedidoId), isNull(pedidos.deletedAt)))
      .limit(1);
    const pr = rows[0];
    if (!pr) throw new Error('Pedido no encontrado.');

    const [lineasRows, historialRows] = await Promise.all([
      tx.select().from(pedidoLineas).where(eq(pedidoLineas.pedidoId, pedidoId)),
      tx
        .select()
        .from(pedidoHistorial)
        .where(eq(pedidoHistorial.pedidoId, pedidoId))
        .orderBy(desc(pedidoHistorial.createdAt))
    ]);

    const actual = mapDbRowsToPedidos([pr], lineasRows, historialRows)[0]!;
    const actualizado = applyPatchToPedido(actual, patch, session.nombre, rolLabel, motivo);

    // Actualización parcial atómica: solo escribir columnas financieras/estado derivadas + campos incluidos en patch
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
          `Acción en lote: marcado como Pagado`
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
