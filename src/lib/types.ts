export type MedioPago = 'Transferencia / Efectivo' | 'Cheque' | 'Sin pagar';
export type EstadoPago = 'Pagado' | 'Pendiente' | 'Falta pago';
export type EstadoArmado = 'Por armar' | 'En armado' | 'Listo' | 'Falta talle o color';
export type EstadoEnvio = 'Pendiente' | 'Despachado' | 'Entregado';
export type EstadoLinea = 'OK' | 'REQUIERE_REVISION';
export type AlertaCheque = 'VENCIDO' | 'VENCE_HOY' | 'EN_PLAZO' | null;
export type VistaActiva = 'pedidos' | 'nuevo';
export type DrawerFocusTarget = 'talles' | 'transporte' | 'derivacion' | null;
export interface ClienteConocido {
  cliente: string;
  vendedor: string;
  whatsapp: string;
}
export type FiltroPedidos =
  | 'todos'
  | 'pendientes_armado'
  | 'pagados'
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
  c?: Array<{ a?: string; t?: string }>;
}

export interface SheetJSMergeRange {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

export interface SheetJSWorkSheet {
  [address: string]:
    | SheetJSCellObject
    | SheetJSMergeRange[]
    | { wch?: number }[]
    | Array<{ hidden?: boolean } | undefined>
    | string
    | undefined;
  '!merges'?: SheetJSMergeRange[];
  '!cols'?: { wch?: number }[];
  '!rows'?: Array<{ hidden?: boolean } | undefined>;
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

export type TipoAdvertenciaExcel =
  | 'CODIGO_DESCONOCIDO'
  | 'CANTIDAD_INVALIDA'
  | 'FILA_SIN_CODIGO'
  | 'FILA_OCULTA'
  | 'FILA_TACHADA'
  | 'NOTA_CONTRADICTORIA'
  | 'ERROR_FORMULA'
  | 'DESCUADRE_FILA';

export type TipoArchivoDetectado =
  | 'LISTA_006'
  | 'PLANILLA_PERSONALIZADA'
  | 'CSV'
  | 'FACTURA_AFIP_PDF'
  | 'PEDIDO_PDF'
  | 'ALTA_CLIENTE_DOCX'
  | 'PEDIDO_DOCX'
  | 'IMAGEN_O_PDF_ESCANEADO'
  | 'PLANILLA_GENERAL_ADMIN'
  | 'COMPROBANTE_NO_PEDIDO'
  | 'NO_COMPATIBLE';

export type EstadoIntegridad = 'EXACTO' | 'REVISAR' | 'BLOQUEADO';

export interface ComentarioCeldaExcel {
  id: string;
  hoja: string;
  fila: number;
  columna: number;
  celdaRef: string;
  codigo: string;
  productoNombre: string;
  texto: string;
  aplicadoComoNota: boolean;
}

export interface DuplicadoEnArchivo {
  codigo: string;
  nombreProducto: string;
  idGrupo: string;
  apariciones: ReadonlyArray<{ fila: number; cantidad: number }>;
  cantidadTotalSumada: number;
}

export interface AuditoriaLectura {
  unidadesLeidas: number;
  unidadesDeclaradasArchivo: number | null;
  diferenciaUnidades: number | null;
  unidadesEnAlertas: number;
  subtotalNetoCalculado: number;
  totalDeclaradoArchivo: number | null;
  diferenciaImporte: number | null;
  filasOmitidasConDatos: number;
  estadoIntegridad: EstadoIntegridad;
  motivos: string[];
  descuentoDeclaradoPorc?: number | null;
  descuentoDeclaradoMonto?: number | null;
  ivaDeclaradoArchivo?: number | null;
  totalFacturadoDeclaradoArchivo?: number | null;
  duplicadosEnArchivo: ReadonlyArray<DuplicadoEnArchivo>;
}

export interface CorreccionAutomatica {
  id: string;
  hoja: string;
  fila: number;
  original: string;
  corregidoA: string;
  productoNombre: string;
  motivo: string;
  lineaId: string;
}

export interface HojaDisponible {
  nombre: string;
  lineas: number;
  unidades: number;
  seleccionada: boolean;
}

export interface DatosClienteDetectados {
  cliente?: string;
  cuit?: string;
  whatsapp?: string;
  email?: string;
  transporte?: string;
  descuentoPorc?: number;
  fechaPedido?: string;
  numeroComprobante?: string;
  tipoComprobante?: string;
  domicilio?: string;
  domicilioEntrega?: string;
  horarioRecepcion?: string;
  condicionIva?: string;
  condicionPago?: string;
  contacto?: string;
  observacionesSugeridas?: string;
  origen: 'ARCHIVO' | 'NOMBRE_ARCHIVO' | 'FACTURA' | 'FICHA_ALTA' | 'IMAGEN';
}

export interface SugerenciaProductoCatalogo {
  idGrupo: string;
  nombre: string;
  codigos: string[];
  rubro: string;
  motivo: string;
}

export interface AvisoPrecioExcel {
  codigo: string;
  nombreProducto: string;
  precioPlanilla: number;
  precioVigente: number;
}

export interface AdvertenciaExcel {
  hoja: string;
  fila: number;
  codigo: string;
  mensaje: string;
  id: string;
  tipo: TipoAdvertenciaExcel;
  codigosDetectados: string[];
  descripcionExcel: string;
  precioExcelReferencia: number | null;
  cantidadDetectada: number | null;
  cantidadRaw: string;
  notaDetectada: string;
  productoExistenteIdGrupo: string | null;
  cantidadAlternativa?: number | null;
  sugerenciasCatalogo: ReadonlyArray<SugerenciaProductoCatalogo>;
}

export interface ResultadoParseoLista006 {
  ok: boolean;
  errorFatal?: string;
  lineas: LineaPedido[];
  advertencias: AdvertenciaExcel[];
  avisosPrecio?: ReadonlyArray<AvisoPrecioExcel>;
  tipoArchivo: TipoArchivoDetectado;
  auditoria: AuditoriaLectura;
  correccionesAutomaticas: ReadonlyArray<CorreccionAutomatica>;
  hojasDisponibles: ReadonlyArray<HojaDisponible>;
  datosDetectados: DatosClienteDetectados | null;
  comentariosDetectados: ReadonlyArray<ComentarioCeldaExcel>;
}

export interface ResultadoAnalisisArchivo extends ResultadoParseoLista006 {
  nombreArchivo: string;
  avisos: string[];
  visionDisponible: boolean;
  requiereConfirmacionHumana: boolean;
  tipoImagen?: 'FICHA_CLIENTE' | 'COMPROBANTE_PAGO' | 'PEDIDO_ARTICULOS' | 'OTRO';
}

export interface CampoAutocompletadoImport {
  campo: 'cliente' | 'whatsapp' | 'observaciones' | 'descuentoPorc' | 'fechaPedido';
  valorAnterior: string;
  valorNuevo: string;
}

export interface ControlImportacionBorrador {
  nombreArchivo: string;
  tipoArchivo: TipoArchivoDetectado;
  auditoria: AuditoriaLectura;
  unidadesBaseBorrador: number;
  descuadreConfirmado: boolean;
  correcciones: ReadonlyArray<CorreccionAutomatica>;
  correccionesRevisadas: boolean;
  duplicados: ReadonlyArray<DuplicadoEnArchivo>;
  duplicadosConfirmados: boolean;
  comentarios: ReadonlyArray<ComentarioCeldaExcel>;
  hojasDisponibles: ReadonlyArray<HojaDisponible>;
  camposAutocompletados: CampoAutocompletadoImport[];
  requiereConfirmacionHumana: boolean;
  confirmacionHumanaOk: boolean;
  avisos: string[];
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

export interface CambioArticuloAjuste {
  tipo: 'modificado' | 'quitado' | 'agregado';
  nombreProducto: string;
  detalle?: string;
  cantidadAnterior: number;
  cantidadNueva: number;
}

export interface ResumenAjusteArticulos {
  cambios: CambioArticuloAjuste[];
  totalAnterior: number;
  totalNuevo: number;
  unidadesAnteriores: number;
  unidadesNuevas: number;
  saldoAFavorGenerado: number;
  diferenciaCheque: number;
  avisoRelecturaDeposito: boolean;
  incidenciaResuelta: boolean;
}
