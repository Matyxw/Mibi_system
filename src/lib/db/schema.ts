import 'server-only';

import {
  boolean,
  date,
  integer,
  jsonb,
  numeric,
  pgSequence,
  pgTable,
  text,
  timestamp,
  uuid
} from 'drizzle-orm/pg-core';
import type { DesgloseVariante } from '@/lib/types';

export const pedidosNumeroSeq = pgSequence('pedidos_numero_seq', { startWith: 1 });

export const usuarios = pgTable('usuarios', {
  id: uuid('id').defaultRandom().primaryKey(),
  username: text('username').notNull().unique(),
  nombre: text('nombre').notNull(),
  rol: text('rol').$type<'ADMIN' | 'VENTAS_DEPOSITO'>().notNull(),
  passwordHash: text('password_hash').notNull(),
  activo: boolean('activo').notNull().default(true),
  tokenVersion: integer('token_version').notNull().default(1),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  updatedAt: timestamp('updated_at', { withTimezone: true }).notNull().defaultNow()
});

export const pedidos = pgTable('pedidos', {
  id: uuid('id').defaultRandom().primaryKey(),
  numeroPedido: text('numero_pedido').notNull().unique(),
  cliente: text('cliente').notNull(),
  vendedor: text('vendedor').notNull().default('Administración'),
  whatsapp: text('whatsapp').notNull().default(''),
  fechaPedido: date('fecha_pedido', { mode: 'string' }).notNull(),
  despachoPrevisto: date('despacho_previsto', { mode: 'string' }),
  fechaDespacho: date('fecha_despacho', { mode: 'string' }),
  fechaEntrega: date('fecha_entrega', { mode: 'string' }),
  descuentoPorc: numeric('descuento_porc', { precision: 5, scale: 2 }).notNull().default('0.00'),
  subtotalNeto: numeric('subtotal_neto', { precision: 14, scale: 2 }).notNull().default('0.00'),
  montoDescuento: numeric('monto_descuento', { precision: 14, scale: 2 }).notNull().default('0.00'),
  baseImponible: numeric('base_imponible', { precision: 14, scale: 2 }).notNull().default('0.00'),
  iva21: numeric('iva_21', { precision: 14, scale: 2 }).notNull().default('0.00'),
  totalFacturado: numeric('total_facturado', { precision: 14, scale: 2 }).notNull().default('0.00'),
  totalFacturadoOverride: numeric('total_facturado_override', { precision: 14, scale: 2 }),
  medioPago: text('medio_pago').$type<'Transferencia / Efectivo' | 'Cheque' | 'Sin pagar'>().notNull().default('Sin pagar'),
  montoAbonado: numeric('monto_abonado', { precision: 14, scale: 2 }).notNull().default('0.00'),
  montoCheque: numeric('monto_cheque', { precision: 14, scale: 2 }).notNull().default('0.00'),
  chequeAcreditado: boolean('cheque_acreditado').notNull().default(false),
  fechaRecepcionCheque: date('fecha_recepcion_cheque', { mode: 'string' }),
  fechaAcreditacionCheque: date('fecha_acreditacion_cheque', { mode: 'string' }),
  detalleCheque: text('detalle_cheque').notNull().default(''),
  saldoPendiente: numeric('saldo_pendiente', { precision: 14, scale: 2 }).notNull().default('0.00'),
  deudaSinRespaldo: numeric('deuda_sin_respaldo', { precision: 14, scale: 2 }).notNull().default('0.00'),
  estadoPago: text('estado_pago').$type<'Pagado' | 'Pendiente' | 'Falta pago'>().notNull().default('Falta pago'),
  estadoArmado: text('estado_armado')
    .$type<'Por armar' | 'En armado' | 'Listo' | 'Falta talle o color'>()
    .notNull()
    .default('Por armar'),
  estadoEnvio: text('estado_envio').$type<'Pendiente' | 'Despachado' | 'Entregado'>().notNull().default('Pendiente'),
  transporte: text('transporte').notNull().default(''),
  nroSeguimiento: text('nro_seguimiento').notNull().default(''),
  observaciones: text('observaciones').notNull().default(''),
  detalleManualImportado: text('detalle_manual_importado'),
  version: integer('version').notNull().default(1),
  deletedAt: timestamp('deleted_at', { withTimezone: true }),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  updatedAt: timestamp('updated_at', { withTimezone: true }).notNull().defaultNow()
});

export const pedidoLineas = pgTable('pedido_lineas', {
  id: uuid('id').defaultRandom().primaryKey(),
  pedidoId: uuid('pedido_id')
    .notNull()
    .references(() => pedidos.id, { onDelete: 'cascade' }),
  idLinea: text('id_linea').notNull(),
  idGrupo: text('id_grupo').notNull(),
  nombreProducto: text('nombre_producto').notNull(),
  rubro: text('rubro').notNull(),
  precioUnitarioNeto: numeric('precio_unitario_neto', { precision: 14, scale: 2 }).notNull(),
  cantidadTotal: integer('cantidad_total').notNull(),
  notaCliente: text('nota_cliente').notNull().default(''),
  estado: text('estado').$type<'OK' | 'REQUIERE_REVISION'>().notNull(),
  requiereRevision: boolean('requiere_revision').notNull().default(false),
  motivoRevision: text('motivo_revision'),
  desgloseJson: jsonb('desglose_json').$type<DesgloseVariante[]>().notNull().default([])
});

export const pedidoHistorial = pgTable('pedido_historial', {
  id: uuid('id').defaultRandom().primaryKey(),
  pedidoId: uuid('pedido_id')
    .notNull()
    .references(() => pedidos.id, { onDelete: 'cascade' }),
  usuarioId: uuid('usuario_id'),
  usuarioNombre: text('usuario_nombre').notNull(),
  rol: text('rol').notNull(),
  accion: text('accion').notNull(),
  createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow()
});
