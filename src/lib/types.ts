export type MedioPago = 'Transferencia / Efectivo' | 'Cheque' | 'Sin pagar';
export type EstadoPago = 'Pagado' | 'Pendiente' | 'Falta pago';
export type EstadoArmado = 'Por armar' | 'En armado' | 'Listo' | 'Falta talle o color';
export type EstadoEnvio = 'Pendiente' | 'Despachado' | 'Entregado';
export type EstadoLinea = 'OK' | 'REQUIERE_REVISION';
export type AlertaCheque = 'VENCIDO' | 'VENCE_HOY' | 'EN_PLAZO' | null;
export type VistaActiva = 'pedidos' | 'nuevo';
export type FiltroPedidos =
  | 'todos'
  | 'falta_talle'
  | 'cheques'
  | 'falta_pago'
  | 'listos'
  | 'sin_enviar_deposito'
  | 'con_aviso_deposito'
  | 'urgentes'
  | 'por_armar'
  | 'en_armado'
  | 'despachados';
export type ToneStatus = 'green' | 'amber' | 'red' | 'neutral' | 'ink';
export type BatchPanelActivo = 'estado' | 'pago' | 'enviar' | 'exportar' | 'eliminar' | null;
export type RolUsuario = 'ADMIN' | 'VENTAS' | 'DEPOSITO' | 'VENTAS_DEPOSITO';
export type RolEquipo = 'Administración' | 'Ventas' | 'Depósito' | 'Ventas / Depósito';
export type PrioridadAsignacion = 'NORMAL' | 'URGENTE';

export const rolUsuarioToLabel = (rol: RolUsuario): RolEquipo => {
  if (rol === 'ADMIN') return 'Administración';
  if (rol === 'VENTAS') return 'Ventas';
  if (rol === 'DEPOSITO') return 'Depósito';
  return 'Ventas / Depósito';
};

export const isDepositoRole = (rol: RolUsuario): boolean => rol === 'DEPOSITO';
export const isVentasRole = (rol: RolUsuario): boolean => rol === 'VENTAS' || rol === 'VENTAS_DEPOSITO';

export type CellPrimitive = string | number | boolean | null;

export interface SheetJSCellObject {
  v?: string | number | boolean | Date;
  t?: string;
  z?: string;
  w?: string;
}

export interface SheetJSMergeRange {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

export interface SheetJSWorkSheet {
  [address: string]: SheetJSCellObject | SheetJSMergeRange[] | { wch?: number }[] | string | undefined;
  '!merges'?: SheetJSMergeRange[];
  '!cols'?: { wch?: number }[];
  '!ref'?: string;
}

export interface SheetJSWorkBook {
  SheetNames: string[];
  Sheets: Record<string, SheetJSWorkSheet>;
}

export interface VarianteCatalogo {
  codigo: string;
  etiqueta: string;
  aliases: string[];
}

export interface ProductoCatalogo {
  idGrupo: string;
  nombre: string;
  rubro: string;
  precioNeto: number;
  variantes: ReadonlyArray<VarianteCatalogo>;
  esCustom?: boolean;
  activo?: boolean;
  customId?: string;
}

export interface ProductoCustomRecord {
  id: string;
  idGrupo: string;
  nombre: string;
  rubro: string;
  precioNeto: number;
  variantes: VarianteCatalogo[];
  activo: boolean;
  createdBy: string | null;
  createdAt: number;
  updatedAt: number;
}

export interface PedidoAsignacion {
  id: string;
  pedidoId: string;
  usuarioId: string;
  usuarioNombre: string;
  usuarioRol: RolUsuario;
  enviadoPor: string | null;
  enviadoPorNombre: string;
  notaEnvio: string;
  prioridad: PrioridadAsignacion;
  vistoAt: number | null;
  requiereRelectura: boolean;
  createdAt: number;
  updatedAt: number;
}

export interface DesgloseVariante {
  codigo: string;
  etiqueta: string;
  cantidad: number;
}

export interface LineaPedido {
  idLinea: string;
  idGrupo: string;
  nombreProducto: string;
  rubro: string;
  precioUnitarioNeto: number;
  cantidadTotal: number;
  notaCliente: string;
  estado: EstadoLinea;
  requiereRevision: boolean;
  motivoRevision?: string;
  desglose: DesgloseVariante[];
}

export interface HistorialEvento {
  fechaHora: string;
  usuario: string;
  rol: RolEquipo;
  accion: string;
}

export interface PedidoBase {
  id: string;
  numeroPedido: string;
  cliente: string;
  vendedor: string;
  whatsapp: string;
  fechaPedido: string;
  despachoPrevisto: string;
  lineas: LineaPedido[];
  descuentoPorc: number;
  medioPago: MedioPago;
  montoAbonado: number;
  montoCheque: number;
  chequeAcreditado: boolean;
  fechaRecepcionCheque: string;
  fechaAcreditacionCheque: string;
  detalleCheque: string;
  estadoArmado: EstadoArmado;
  estadoEnvio: EstadoEnvio;
  fechaDespacho?: string;
  fechaEntrega?: string;
  transporte: string;
  nroSeguimiento: string;
  observaciones: string;
  version: number;
  deletedAt?: string | null;
  createdAt: number;
  updatedAt: number;
  historial: HistorialEvento[];
  createdBy?: string | null;
  incidenciaDeposito?: string;
  incidenciaAt?: number | null;
  incidenciaPor?: string;
  asignaciones?: PedidoAsignacion[];
  detalleManualImportado?: string;
  totalFacturadoOverride?: number;
}

export interface Pedido extends PedidoBase {
  subtotalNeto: number;
  montoDescuento: number;
  baseImponible: number;
  iva21: number;
  totalFacturado: number;
  saldoPendiente: number;
  deudaSinRespaldo: number;
  estadoPago: EstadoPago;
  faltantesVariante: number;
  detalleResumen: string;
  incidenciaDeposito: string;
  incidenciaAt: number | null;
  incidenciaPor: string;
  asignaciones: PedidoAsignacion[];
}

export interface ResumenFinanciero {
  porArmar: number;
  cobradoEnMano: number;
  enCheques: number;
  chequesPorVencer: number;
  faltaPago: number;
}

export interface AdvertenciaExcel {
  hoja: string;
  fila: number;
  codigo: string;
  mensaje: string;
}

export interface ResultadoParseoLista006 {
  ok: boolean;
  errorFatal?: string;
  lineas: LineaPedido[];
  advertencias: AdvertenciaExcel[];
}

export interface ResultadoPickingConsolidado {
  success: boolean;
  incluidos: number;
  excluidosTalle: string[];
  excluidosSinLineas: string[];
  totalUnidades: number;
  error?: string;
}

export interface SessionUser {
  userId: string;
  username: string;
  nombre: string;
  rol: RolUsuario;
  tokenVersion: number;
}

export interface UsuarioEquipo {
  id: string;
  username: string;
  nombre: string;
  rol: RolUsuario;
  activo: boolean;
  tokenVersion: number;
  createdAt: number;
}

export interface ResultadoQA {
  nombre: string;
  ok: boolean;
  detalle: string;
}

export interface ToastState {
  texto: string;
  tone: ToneStatus;
}
