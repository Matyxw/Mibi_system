'use client';

import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Search,
  Plus,
  Upload,
  Download,
  ChevronRight,
  X,
  Check,
  Copy,
  Printer,
  Trash2,
  Lock,
  LogOut,
  Users,
  RotateCcw,
  FileSpreadsheet,
  Clock,
  AlertTriangle,
  CheckCircle2,
  Truck,
  CreditCard,
  Layers,
  Undo2,
  Send,
  MessageCircle,
  Package,
  ShieldCheck,
  ArrowLeft
} from 'lucide-react';
import { z } from 'zod';
import JSZip from 'jszip';
import { initializeApp, getApps, type FirebaseApp } from 'firebase/app';
import { getAuth, signInAnonymously, onAuthStateChanged, type User as FirebaseUser } from 'firebase/auth';
import {
  getFirestore,
  doc,
  setDoc,
  updateDoc,
  onSnapshot,
  writeBatch,
  collection,
  type Firestore,
  type DocumentReference
} from 'firebase/firestore';

// ============================================================================
// IDENTIDAD OFICIAL MIBI® & LOGOTIPO VECTORIAL
// ============================================================================

const MIBI_BRAND = {
  nombre: 'Mibi®',
  pieOficial: 'Mibi® · Ciudad de la Paz 2467, Belgrano, CABA · WhatsApp 11-5606-9384 · contacto@mibi.com.ar',
  web: 'https://www.mibi.com.ar/'
} as const;

const VENDEDORES_MIBI = ['Juancho', 'Pablo Palacios', 'Agustín', 'Rocío', 'Administración'] as const;

type LogoVariant = 'ink' | 'pink' | 'white';

function MibiLogo({
  width = 88,
  variant = 'ink'
}: {
  width?: number;
  variant?: LogoVariant;
}): React.JSX.Element {
  const src =
    variant === 'pink'
      ? '/mibi-logo-pink.png'
      : variant === 'white'
        ? '/mibi-logo-white.png'
        : '/mibi-logo-ink.png';
  return (
    <img
      src={src}
      alt="Mibi®"
      width={width}
      height={Math.round(width * (210 / 480))}
      style={{ width, height: 'auto' }}
      className="block select-none object-contain"
      draggable={false}
    />
  );
}

// ============================================================================
// TIPOS DE DOMINIO (STRICT TYPESCRIPT — CERO ANY)
// ============================================================================

type MedioPago = 'Transferencia / Efectivo' | 'Cheque' | 'Sin pagar';
type EstadoPago = 'Pagado' | 'Pendiente' | 'Falta pago';
type EstadoArmado = 'Por armar' | 'En armado' | 'Listo' | 'Falta talle o color';
type EstadoEnvio = 'Pendiente' | 'Despachado' | 'Entregado';
type EstadoLinea = 'OK' | 'REQUIERE_REVISION';
type AlertaCheque = 'VENCIDO' | 'VENCE_HOY' | 'EN_PLAZO' | null;
type VistaActiva = 'pedidos' | 'nuevo';
type FiltroPedidos = 'todos' | 'falta_talle' | 'cheques' | 'falta_pago' | 'listos';
type ToneStatus = 'green' | 'amber' | 'red' | 'neutral' | 'ink';
type BatchPanelActivo = 'estado' | 'pago' | 'exportar' | 'eliminar' | null;
type RolEquipo = 'Administración' | 'Ventas / Depósito';

type CellPrimitive = string | number | boolean | null;

interface SheetJSCellObject {
  v?: string | number | boolean | Date;
  t?: string;
  z?: string;
  w?: string;
}

interface SheetJSMergeRange {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

interface SheetJSWorkSheet {
  [address: string]: SheetJSCellObject | SheetJSMergeRange[] | { wch?: number }[] | string | undefined;
  '!merges'?: SheetJSMergeRange[];
  '!cols'?: { wch?: number }[];
  '!ref'?: string;
}

interface SheetJSWorkBook {
  SheetNames: string[];
  Sheets: Record<string, SheetJSWorkSheet>;
}

interface SheetJSModule {
  read: (
    data: ArrayBuffer | Uint8Array | string,
    opts?: { type?: 'array' | 'binary' | 'string'; raw?: boolean; codepage?: number }
  ) => SheetJSWorkBook;
  writeFile: (wb: SheetJSWorkBook, filename: string) => void;
  utils: {
    book_new: () => SheetJSWorkBook;
    book_append_sheet: (wb: SheetJSWorkBook, ws: SheetJSWorkSheet, name: string) => void;
    aoa_to_sheet: (data: CellPrimitive[][]) => SheetJSWorkSheet;
    sheet_to_json: <T = CellPrimitive[]>(
      ws: SheetJSWorkSheet,
      opts?: { header?: number | string[]; defval?: CellPrimitive; raw?: boolean }
    ) => T[];
    encode_cell: (cell: { r: number; c: number }) => string;
    decode_cell: (addr: string) => { r: number; c: number };
    decode_range: (ref: string) => SheetJSMergeRange;
  };
}

interface VarianteCatalogo {
  codigo: string;
  etiqueta: string;
  aliases: string[];
}

interface ProductoCatalogo {
  idGrupo: string;
  nombre: string;
  rubro: string;
  precioNeto: number;
  variantes: ReadonlyArray<VarianteCatalogo>;
}

interface DesgloseVariante {
  codigo: string;
  etiqueta: string;
  cantidad: number;
}

interface LineaPedido {
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

interface HistorialEvento {
  fechaHora: string;
  usuario: string;
  rol: RolEquipo;
  accion: string;
}

interface PedidoBase {
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
  createdAt: number;
  updatedAt: number;
  historial: HistorialEvento[];
  detalleManualImportado?: string;
  totalFacturadoOverride?: number;
}

interface Pedido extends PedidoBase {
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
}

interface ResumenFinanciero {
  porArmar: number;
  cobradoEnMano: number;
  enCheques: number;
  chequesPorVencer: number;
  faltaPago: number;
}

interface PendingUndoDelete {
  pedidos: Pedido[];
  timerId: number;
}

interface ResultadoPickingConsolidado {
  success: boolean;
  incluidos: number;
  excluidosTalle: string[];
  excluidosSinLineas: string[];
  totalUnidades: number;
  error?: string;
}

interface ResultadoHojaDeposito {
  success: boolean;
  errors: string[];
}

interface AdvertenciaExcel {
  hoja: string;
  fila: number;
  codigo: string;
  mensaje: string;
}

interface ResultadoParseoLista006 {
  ok: boolean;
  errorFatal?: string;
  lineas: LineaPedido[];
  advertencias: AdvertenciaExcel[];
}

interface ConfigEquipo {
  adminPinHash: string;
  ventasPinHash: string;
  updatedAt: number;
  updatedBy: string;
}

interface SesionUsuario {
  nombre: string;
  rol: RolEquipo;
  configVersion: number;
}

interface ResultadoQA {
  nombre: string;
  ok: boolean;
  detalle: string;
}

interface ToastState {
  texto: string;
  tone: ToneStatus;
}

interface FormNuevoPedido {
  cliente: string;
  vendedor: string;
  whatsapp: string;
  fechaPedido: string;
  despachoPrevisto: string;
  observaciones: string;
  descuentoPorc: number;
  pagoTipo: MedioPago;
  montoAbonadoInput: string;
  montoChequeInput: string;
  fechaRecepcionCheque: string;
  fechaAcreditacionCheque: string;
  detalleCheque: string;
}

interface DrawerDraft {
  cliente: string;
  vendedor: string;
  whatsapp: string;
  despachoPrevisto: string;
  descuentoPorc: string;
  montoAbonado: string;
  montoCheque: string;
  fechaRecepcionCheque: string;
  fechaAcreditacionCheque: string;
  detalleCheque: string;
  transporte: string;
  nroSeguimiento: string;
  observaciones: string;
}

// ============================================================================
// UTILIDADES DE REDONDEO EXACTO, FECHAS Y FORMATO
// ============================================================================

const round2 = (val: number): number => {
  if (!Number.isFinite(val)) return 0;
  const sign = val < 0 ? -1 : 1;
  const abs = Math.abs(val);
  return (sign * Math.round(Number((abs * 100).toFixed(6)) + 1e-7)) / 100;
};

const getTodayISO = (): string => {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

const addDaysISO = (isoDate: string, days: number): string => {
  const base = /^\d{4}-\d{2}-\d{2}$/.test(isoDate) ? isoDate : getTodayISO();
  const dt = new Date(`${base}T12:00:00`);
  dt.setDate(dt.getDate() + days);
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  const d = String(dt.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

const formatShortDate = (isoDate: string): string => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return isoDate || '—';
  const [, m, d] = isoDate.split('-');
  return `${d}/${m}`;
};

const formatNowDateTime = (): string => {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  return `${dd}/${mm} ${hh}:${min}`;
};

const formatARS = (val: number): string =>
  new Intl.NumberFormat('es-AR', {
    style: 'currency',
    currency: 'ARS',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(round2(val));

const formatARSCompacto = (val: number): string => {
  const r = round2(val);
  const sinCentavos = Math.abs(r - Math.round(r)) < 0.005;
  return new Intl.NumberFormat('es-AR', {
    style: 'currency',
    currency: 'ARS',
    minimumFractionDigits: sinCentavos ? 0 : 2,
    maximumFractionDigits: sinCentavos ? 0 : 2
  }).format(r);
};

const readString = (obj: Record<string, unknown>, key: string, fallback = ''): string => {
  const val = obj[key];
  return typeof val === 'string' ? val : fallback;
};

const readNumber = (obj: Record<string, unknown>, key: string, fallback = 0): number => {
  const val = obj[key];
  return typeof val === 'number' && Number.isFinite(val) ? val : fallback;
};

const readBoolean = (obj: Record<string, unknown>, key: string, fallback = false): boolean => {
  const val = obj[key];
  return typeof val === 'boolean' ? val : fallback;
};

const copyTextSafe = (text: string): void => {
  if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
    void navigator.clipboard.writeText(text).catch(() => {
      // Fallback silencioso
    });
    return;
  }
  if (typeof document !== 'undefined') {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand('copy');
    } catch {
      // Ignorar si está bloqueado
    }
    document.body.removeChild(ta);
  }
};

const hashPin = async (pin: string): Promise<string> => {
  const normalized = pin.trim();
  if (typeof window !== 'undefined' && window.crypto?.subtle) {
    const data = new TextEncoder().encode(`mibi_erp_v1:${normalized}`);
    const digest = await window.crypto.subtle.digest('SHA-256', data);
    return Array.from(new Uint8Array(digest))
      .map((b) => b.toString(16).padStart(2, '0'))
      .join('');
  }
  let h = 2166136261;
  for (let i = 0; i < normalized.length; i++) {
    h ^= normalized.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return `fnv_${(h >>> 0).toString(16)}`;
};

// ============================================================================
// CATÁLOGO OFICIAL INMUTABLE — LISTA N° 006 (CÓDIGOS 6500–6614)
// ============================================================================

const CATALOGO_MIBI: ReadonlyArray<ProductoCatalogo> = [
  // 1. CHALECOS Y FLOTADORES
  {
    idGrupo: '6508-6509',
    nombre: 'Flotador Bracitos Monster (11 a 19 kg)',
    rubro: 'FLOTA',
    precioNeto: 23800,
    variantes: [
      { codigo: '6508', etiqueta: 'Green (Verde)', aliases: ['g', 'green', 'verde', '6508'] },
      { codigo: '6509', etiqueta: 'Pink (Rosa)', aliases: ['r', 'pink', 'rosa', '6509'] }
    ]
  },
  {
    idGrupo: '6500-6501',
    nombre: 'Chaleco Bubble Flotación Retirable (15 a 20 kg)',
    rubro: 'CH',
    precioNeto: 40661,
    variantes: [
      { codigo: '6500', etiqueta: 'Azul / Aqua', aliases: ['a', 'azul', 'aqua', 'celeste', '6500'] },
      { codigo: '6501', etiqueta: 'Pink (Rosa)', aliases: ['r', 'pink', 'rosa', '6501'] }
    ]
  },
  {
    idGrupo: '6502-6503',
    nombre: 'Chaleco Monster Placa Completa (3 a 6 años)',
    rubro: 'CH',
    precioNeto: 42645,
    variantes: [
      { codigo: '6502', etiqueta: 'Green (Verde)', aliases: ['g', 'green', 'verde', '6502'] },
      { codigo: '6503', etiqueta: 'Pink (Rosa)', aliases: ['r', 'pink', 'rosa', '6503'] }
    ]
  },
  {
    idGrupo: '6504-6505',
    nombre: 'Chaleco Shark Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6504', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6504'] },
      { codigo: '6505', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6505'] }
    ]
  },
  {
    idGrupo: '6562-6563',
    nombre: 'Chaleco Coco Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6562', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6562'] },
      { codigo: '6563', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6563'] }
    ]
  },
  {
    idGrupo: '6564-6565',
    nombre: 'Chaleco Rabb Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6564', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6564'] },
      { codigo: '6565', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6565'] }
    ]
  },
  {
    idGrupo: '6566-6567',
    nombre: 'Chaleco Lovely Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6566', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6566'] },
      { codigo: '6567', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6567'] }
    ]
  },
  {
    idGrupo: '6540-6568',
    nombre: 'Chaleco Bug Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6540', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6540'] },
      { codigo: '6568', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6568'] }
    ]
  },
  {
    idGrupo: '6541-6569',
    nombre: 'Chaleco Margarita Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6541', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6541'] },
      { codigo: '6569', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6569'] }
    ]
  },
  {
    idGrupo: '6539',
    nombre: 'Chaleco Petit Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [{ codigo: '6539', etiqueta: 'Único (1-3 años)', aliases: ['unico', '6539'] }]
  },
  {
    idGrupo: '6542',
    nombre: 'Chaleco Ocean Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [{ codigo: '6542', etiqueta: 'Único (1-3 años)', aliases: ['unico', '6542'] }]
  },
  {
    idGrupo: '6543',
    nombre: 'Chaleco Lila Splash Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [{ codigo: '6543', etiqueta: 'Único (1-3 años)', aliases: ['unico', '6543'] }]
  },
  {
    idGrupo: '6545',
    nombre: 'Chaleco Crabby Cool Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [{ codigo: '6545', etiqueta: 'Único (1-3 años)', aliases: ['unico', '6545'] }]
  },

  // 2. ANTIPARRAS Y MÁSCARAS
  {
    idGrupo: '6510',
    nombre: 'Antiparras Dino',
    rubro: 'ANTI',
    precioNeto: 9505,
    variantes: [{ codigo: '6510', etiqueta: 'Único Kids', aliases: ['unico', '6510'] }]
  },
  {
    idGrupo: '6511',
    nombre: 'Antiparras Pop',
    rubro: 'ANTI',
    precioNeto: 9505,
    variantes: [{ codigo: '6511', etiqueta: 'Único Kids', aliases: ['unico', '6511'] }]
  },
  {
    idGrupo: '6512',
    nombre: 'Antiparras Unicorn',
    rubro: 'ANTI',
    precioNeto: 10744,
    variantes: [{ codigo: '6512', etiqueta: 'Único Kids', aliases: ['unico', '6512'] }]
  },
  {
    idGrupo: '6513',
    nombre: 'Antiparras Sweet',
    rubro: 'ANTI',
    precioNeto: 10744,
    variantes: [{ codigo: '6513', etiqueta: 'Único Kids', aliases: ['unico', '6513'] }]
  },
  {
    idGrupo: '6528',
    nombre: 'Antiparras Candy',
    rubro: 'ANTI',
    precioNeto: 14050,
    variantes: [{ codigo: '6528', etiqueta: 'Único Kids', aliases: ['unico', '6528'] }]
  },
  {
    idGrupo: '6530',
    nombre: 'Antiparras Gummy',
    rubro: 'ANTI',
    precioNeto: 14050,
    variantes: [{ codigo: '6530', etiqueta: 'Único Kids', aliases: ['unico', '6530'] }]
  },
  {
    idGrupo: '6536',
    nombre: 'Antiparras Siren',
    rubro: 'ANTI',
    precioNeto: 17850,
    variantes: [{ codigo: '6536', etiqueta: 'Único Kids', aliases: ['unico', '6536'] }]
  },
  {
    idGrupo: '6537',
    nombre: 'Antiparras Shark',
    rubro: 'ANTI',
    precioNeto: 17850,
    variantes: [{ codigo: '6537', etiqueta: 'Único Kids', aliases: ['unico', '6537'] }]
  },
  {
    idGrupo: '6594',
    nombre: 'Máscara Snorkel Hipo',
    rubro: 'MASCA',
    precioNeto: 28100,
    variantes: [{ codigo: '6594', etiqueta: 'Único Kids', aliases: ['unico', '6594'] }]
  },
  {
    idGrupo: '6595',
    nombre: 'Máscara Snorkel Sharky',
    rubro: 'MASCA',
    precioNeto: 28100,
    variantes: [{ codigo: '6595', etiqueta: 'Único Kids', aliases: ['unico', '6595'] }]
  },

  // 3. ACCESORIOS DE AGUA (PAÑALES, ZAPATILLAS, MOCHILAS)
  {
    idGrupo: '6570',
    nombre: 'Pañal de Agua Shark (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6570', etiqueta: 'Único (12-24m)', aliases: ['unico', '6570'] }]
  },
  {
    idGrupo: '6571',
    nombre: 'Pañal de Agua Rabb (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6571', etiqueta: 'Único (12-24m)', aliases: ['unico', '6571'] }]
  },
  {
    idGrupo: '6572',
    nombre: 'Pañal de Agua Margarita (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6572', etiqueta: 'Único (12-24m)', aliases: ['unico', '6572'] }]
  },
  {
    idGrupo: '6573',
    nombre: 'Pañal de Agua Coco (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6573', etiqueta: 'Único (12-24m)', aliases: ['unico', '6573'] }]
  },
  {
    idGrupo: '6574',
    nombre: 'Pañal de Agua Bug (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6574', etiqueta: 'Único (12-24m)', aliases: ['unico', '6574'] }]
  },
  {
    idGrupo: '6575',
    nombre: 'Pañal de Agua Lovely (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6575', etiqueta: 'Único (12-24m)', aliases: ['unico', '6575'] }]
  },
  {
    idGrupo: '6576-6577',
    nombre: 'Zapatillas de Agua Shark',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6576', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6576'] },
      { codigo: '6577', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6577'] }
    ]
  },
  {
    idGrupo: '6578-6579',
    nombre: 'Zapatillas de Agua Rabb',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6578', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6578'] },
      { codigo: '6579', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6579'] }
    ]
  },
  {
    idGrupo: '6580-6581',
    nombre: 'Zapatillas de Agua Margarita',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6580', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6580'] },
      { codigo: '6581', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6581'] }
    ]
  },
  {
    idGrupo: '6582-6583',
    nombre: 'Zapatillas de Agua Coco',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6582', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6582'] },
      { codigo: '6583', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6583'] }
    ]
  },
  {
    idGrupo: '6584-6585',
    nombre: 'Zapatillas de Agua Bug',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6584', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6584'] },
      { codigo: '6585', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6585'] }
    ]
  },
  {
    idGrupo: '6586-6587',
    nombre: 'Zapatillas de Agua Lovely',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6586', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6586'] },
      { codigo: '6587', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6587'] }
    ]
  },
  {
    idGrupo: '6588',
    nombre: 'Mochila de Arena Shark',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6588', etiqueta: 'Único Kids', aliases: ['unico', '6588'] }]
  },
  {
    idGrupo: '6589',
    nombre: 'Mochila de Arena Rabb',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6589', etiqueta: 'Único Kids', aliases: ['unico', '6589'] }]
  },
  {
    idGrupo: '6590',
    nombre: 'Mochila de Arena Coco',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6590', etiqueta: 'Único Kids', aliases: ['unico', '6590'] }]
  },
  {
    idGrupo: '6591',
    nombre: 'Mochila de Arena Margarita',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6591', etiqueta: 'Único Kids', aliases: ['unico', '6591'] }]
  },
  {
    idGrupo: '6592',
    nombre: 'Mochila de Arena Bug',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6592', etiqueta: 'Único Kids', aliases: ['unico', '6592'] }]
  },
  {
    idGrupo: '6593',
    nombre: 'Mochila de Arena Lovely',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6593', etiqueta: 'Único Kids', aliases: ['unico', '6593'] }]
  },

  // 4. INFLABLES Y FLOTADORES
  {
    idGrupo: '6602',
    nombre: 'Inflable Sunny Pink',
    rubro: 'INFLA',
    precioNeto: 30580,
    variantes: [{ codigo: '6602', etiqueta: 'Único', aliases: ['unico', '6602'] }]
  },
  {
    idGrupo: '6603',
    nombre: 'Inflable Sunny Aqua',
    rubro: 'INFLA',
    precioNeto: 30580,
    variantes: [{ codigo: '6603', etiqueta: 'Único', aliases: ['unico', '6603'] }]
  },
  {
    idGrupo: '6604',
    nombre: 'Inflable Wave Pink',
    rubro: 'INFLA',
    precioNeto: 27275,
    variantes: [{ codigo: '6604', etiqueta: 'Único', aliases: ['unico', '6604'] }]
  },
  {
    idGrupo: '6605',
    nombre: 'Inflable Wave Blue',
    rubro: 'INFLA',
    precioNeto: 27275,
    variantes: [{ codigo: '6605', etiqueta: 'Único', aliases: ['unico', '6605'] }]
  },
  {
    idGrupo: '6606',
    nombre: 'Flotador Bunny Hug',
    rubro: 'FLOTA',
    precioNeto: 37190,
    variantes: [{ codigo: '6606', etiqueta: 'Único', aliases: ['unico', '6606'] }]
  },
  {
    idGrupo: '6607',
    nombre: 'Flotador Dino Hug',
    rubro: 'FLOTA',
    precioNeto: 37190,
    variantes: [{ codigo: '6607', etiqueta: 'Único', aliases: ['unico', '6607'] }]
  },
  {
    idGrupo: '6608',
    nombre: 'Pileta Love Pool',
    rubro: 'INFLA',
    precioNeto: 28925,
    variantes: [{ codigo: '6608', etiqueta: 'Único', aliases: ['unico', '6608'] }]
  },
  {
    idGrupo: '6609',
    nombre: 'Pileta Baby Dino',
    rubro: 'INFLA',
    precioNeto: 29752,
    variantes: [{ codigo: '6609', etiqueta: 'Único', aliases: ['unico', '6609'] }]
  },

  // 5. VARIOS Y PACKAGING
  {
    idGrupo: '6610',
    nombre: 'Pelela Lola V1',
    rubro: 'PELELAS',
    precioNeto: 33060,
    variantes: [{ codigo: '6610', etiqueta: 'Único', aliases: ['unico', '6610'] }]
  },
  {
    idGrupo: '6611',
    nombre: 'Pelela Lola V2',
    rubro: 'PELELAS',
    precioNeto: 33060,
    variantes: [{ codigo: '6611', etiqueta: 'Único', aliases: ['unico', '6611'] }]
  },
  {
    idGrupo: '6612',
    nombre: 'Bolsa Troquelada 33x35',
    rubro: 'PACK',
    precioNeto: 3000,
    variantes: [{ codigo: '6612', etiqueta: 'Único', aliases: ['unico', '6612'] }]
  },
  {
    idGrupo: '6613',
    nombre: 'Bolso Playero con Fuelle',
    rubro: 'PACK',
    precioNeto: 7000,
    variantes: [{ codigo: '6613', etiqueta: 'Único', aliases: ['unico', '6613'] }]
  },
  {
    idGrupo: '6614',
    nombre: 'Bolsa PVC 28x28',
    rubro: 'PACK',
    precioNeto: 3500,
    variantes: [{ codigo: '6614', etiqueta: 'Único', aliases: ['unico', '6614'] }]
  }
];

const CATALOGO_BY_GRUPO = new Map<string, ProductoCatalogo>();
const CATALOGO_BY_CODIGO = new Map<string, ProductoCatalogo>();

for (const prod of CATALOGO_MIBI) {
  CATALOGO_BY_GRUPO.set(prod.idGrupo, prod);
  for (const v of prod.variantes) {
    CATALOGO_BY_CODIGO.set(v.codigo, prod);
  }
}

// ============================================================================
// MOTOR FINANCIERO ("PLATA EN MANO") Y REGLAS DE ESTADO
// ============================================================================

const getChequeAlertStatus = (pedido: Pedido): AlertaCheque => {
  if (pedido.medioPago !== 'Cheque' || pedido.chequeAcreditado || pedido.montoCheque <= 0) {
    return null;
  }
  const hoy = getTodayISO();
  const venc = pedido.fechaAcreditacionCheque;
  if (!venc) return null;
  if (venc < hoy) return 'VENCIDO';
  if (venc === hoy) return 'VENCE_HOY';
  return 'EN_PLAZO';
};

// Días calendario restantes hasta la acreditación del cheque (0 = hoy, <0 = vencido)
const getDiasParaAcreditacionCheque = (pedido: Pedido): number | null => {
  if (pedido.medioPago !== 'Cheque' || pedido.chequeAcreditado || pedido.montoCheque <= 0) {
    return null;
  }
  const venc = pedido.fechaAcreditacionCheque;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(venc)) return null;
  const hoyMs = new Date(`${getTodayISO()}T12:00:00`).getTime();
  const vencMs = new Date(`${venc}T12:00:00`).getTime();
  return Math.round((vencMs - hoyMs) / 86400000);
};

const esDespachoAtrasado = (pedido: Pedido): boolean => {
  if (pedido.estadoEnvio !== 'Pendiente') return false;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(pedido.despachoPrevisto)) return false;
  return pedido.despachoPrevisto < getTodayISO();
};

const construirDetalleResumen = (lineas: ReadonlyArray<LineaPedido>, fallback?: string): string => {
  if (lineas.length === 0) return fallback?.trim() || '—';
  const rubros = Array.from(new Set(lineas.map((l) => l.rubro)));
  const totalUnidades = lineas.reduce((acc, l) => acc + l.cantidadTotal, 0);
  return `${rubros.join(' · ')} (${totalUnidades} u.)`;
};

const recalcularPedido = (base: PedidoBase): Pedido => {
  const descuentoPorc = Math.max(0, Math.min(100, Number.isFinite(base.descuentoPorc) ? base.descuentoPorc : 0));

  const subtotalNetoCalculado = round2(
    base.lineas.reduce((acc, l) => acc + round2(l.precioUnitarioNeto * l.cantidadTotal), 0)
  );
  const montoDescuentoCalculado = round2((subtotalNetoCalculado * descuentoPorc) / 100);
  const baseImponibleCalculada = round2(subtotalNetoCalculado - montoDescuentoCalculado);
  const iva21Calculado = round2(baseImponibleCalculada * 0.21);
  const totalFacturadoCalculado = round2(baseImponibleCalculada + iva21Calculado);

  const usarOverride =
    base.lineas.length === 0 &&
    typeof base.totalFacturadoOverride === 'number' &&
    Number.isFinite(base.totalFacturadoOverride) &&
    base.totalFacturadoOverride >= 0;

  const totalFacturado = usarOverride ? round2(base.totalFacturadoOverride ?? 0) : totalFacturadoCalculado;
  const baseImponible = usarOverride ? round2(totalFacturado / 1.21) : baseImponibleCalculada;
  const iva21 = usarOverride ? round2(totalFacturado - baseImponible) : iva21Calculado;
  const subtotalNeto = usarOverride ? baseImponible : subtotalNetoCalculado;
  const montoDescuento = usarOverride ? 0 : montoDescuentoCalculado;

  let montoAbonado = Math.max(0, round2(base.montoAbonado));
  if (montoAbonado > totalFacturado) {
    montoAbonado = totalFacturado;
  }

  let montoCheque =
    base.medioPago === 'Cheque' && !base.chequeAcreditado
      ? Math.max(0, round2(base.montoCheque > 0 ? base.montoCheque : round2(totalFacturado - montoAbonado)))
      : 0;
  if (montoAbonado + montoCheque > totalFacturado) {
    montoCheque = Math.max(0, round2(totalFacturado - montoAbonado));
  }

  const saldoPendiente = Math.max(0, round2(totalFacturado - montoAbonado));
  const deudaSinRespaldo = Math.max(0, round2(totalFacturado - montoAbonado - montoCheque));

  let estadoPago: EstadoPago;
  if (base.medioPago === 'Cheque' && !base.chequeAcreditado) {
    estadoPago = saldoPendiente <= 0.01 ? 'Pagado' : 'Pendiente';
  } else if (saldoPendiente <= 0.01 && totalFacturado > 0) {
    estadoPago = 'Pagado';
  } else if (montoAbonado > 0) {
    estadoPago = 'Pendiente';
  } else {
    estadoPago = 'Falta pago';
  }

  const faltantesVariante = base.lineas.filter((l) => l.requiereRevision || l.estado !== 'OK').length;

  let estadoArmado: EstadoArmado = base.estadoArmado;
  if (faltantesVariante > 0) {
    estadoArmado = 'Falta talle o color';
  } else if (estadoArmado === 'Falta talle o color') {
    estadoArmado = 'Por armar';
  }

  return {
    ...base,
    descuentoPorc,
    montoAbonado,
    montoCheque,
    estadoArmado,
    subtotalNeto,
    montoDescuento,
    baseImponible,
    iva21,
    totalFacturado,
    saldoPendiente,
    deudaSinRespaldo,
    estadoPago,
    faltantesVariante,
    detalleResumen: construirDetalleResumen(base.lineas, base.detalleManualImportado)
  };
};

const calcularResumenFinanciero = (pedidos: ReadonlyArray<Pedido>): ResumenFinanciero => {
  let porArmar = 0;
  let cobradoEnMano = 0;
  let enCheques = 0;
  let chequesPorVencer = 0;
  let faltaPago = 0;

  for (const p of pedidos) {
    if (
      p.estadoArmado === 'Por armar' ||
      p.estadoArmado === 'En armado' ||
      p.estadoArmado === 'Falta talle o color'
    ) {
      porArmar += 1;
    }
    cobradoEnMano = round2(cobradoEnMano + p.montoAbonado);
    if (p.medioPago === 'Cheque' && !p.chequeAcreditado && p.montoCheque > 0) {
      enCheques = round2(enCheques + p.montoCheque);
      const alerta = getChequeAlertStatus(p);
      if (alerta === 'VENCIDO' || alerta === 'VENCE_HOY') chequesPorVencer += 1;
    }
    faltaPago = round2(faltaPago + p.deudaSinRespaldo);
  }

  return {
    porArmar,
    cobradoEnMano: round2(cobradoEnMano),
    enCheques: round2(enCheques),
    chequesPorVencer,
    faltaPago: round2(faltaPago)
  };
};

// ============================================================================
// MOTOR DETERMINÍSTICO DE TALLES/COLORES, PARSER LISTA N° 006 E IMPORTADOR
// ============================================================================

const limpiarTextoInvisible = (s: string): string => s.replace(/[\u200B-\u200D\uFEFF]/g, '').trim();

const normalizarTexto = (s: string): string =>
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

const resolverLinea = (
  producto: ProductoCatalogo,
  cantidadTotal: number,
  notaClienteRaw: string,
  idLineaSuffix = ''
): LineaPedido => {
  const notaLimpia = limpiarTextoInvisible(notaClienteRaw);
  const notaNorm = normalizarTexto(notaLimpia);
  const idLinea = `${producto.idGrupo}${idLineaSuffix ? `-${idLineaSuffix}` : ''}`;

  // Caso 1: Producto de variante única -> asignación directa 100% segura
  if (producto.variantes.length === 1) {
    const v = producto.variantes[0];
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

  // Caso 2: Producto con 2 variantes (talle o color) -> cero suposiciones
  const [v0, v1] = producto.variantes;
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

  // Verificar mención de colores inválidos que no corresponden a las variantes del producto
  const aliasesTodos = new Set([...v0.aliases, ...v1.aliases]);
  for (const colorInv of COLORES_EXTRA_INVALIDOS) {
    const reColor = new RegExp(`\\b${colorInv}\\b`, 'i');
    if (reColor.test(notaNorm) && !aliasesTodos.has(colorInv)) {
      return fallbackRevision(`Color "${colorInv}" no disponible en este modelo`);
    }
  }

  // "mitad y mitad" o "50 y 50"
  if (/\b(mitad y mitad|50 y 50|50\/50)\b/.test(notaNorm)) {
    if (cantidadTotal % 2 !== 0) {
      return fallbackRevision(`Cantidad impar (${cantidadTotal}) para repartir mitad y mitad`);
    }
    const mitad = cantidadTotal / 2;
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
        { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: mitad },
        { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: mitad }
      ]
    };
  }

  // Patrón puro "a y b" / "a / b" / "a + b" (solo números simétricos aceptados automáticamente)
  const matchPuro = notaNorm.match(/^(\d+)\s*(?:y|\/|\+|-)\s*(\d+)$/);
  if (matchPuro) {
    const a = parseInt(matchPuro[1], 10);
    const b = parseInt(matchPuro[2], 10);
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

  // Evaluar menciones explícitas a T1 ("de 1 a 3") o T2 ("de 3 a 6") o colores específicos
  const tieneRangoT1 = /\b(?:de\s*)?1\s*a\s*3\b|\bt1\b|\btalle\s*1\b/.test(notaNorm);
  const tieneRangoT2 = /\b(?:de\s*)?3\s*a\s*6\b|\bt2\b|\btalle\s*2\b/.test(notaNorm);

  // Caso especial auditado: "de 1 a 3: 2" sobre total 4 -> debe frenarse
  const matchRangoConCantidad = notaNorm.match(
    /^(?:de\s*1\s*a\s*3|1\s*a\s*3|t1|talle\s*1)\s*(?::|=|-)?\s*(\d+)$/
  );
  if (matchRangoConCantidad && !tieneRangoT2) {
    const qtyIndicada = parseInt(matchRangoConCantidad[1], 10);
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

  // Nota limpia sin números extra que nombra una sola variante ("de 1 a 3", "todo t1", "verde", "rosa")
  const notaSinRangosEdad = notaNorm
    .replace(/\b(?:de\s*)?1\s*a\s*3(?:\s*anos)?\b/g, ' ')
    .replace(/\b(?:de\s*)?3\s*a\s*6(?:\s*anos)?\b/g, ' ')
    .replace(/\btalle\s*[12]\b/g, ' ')
    .replace(/\bt[12]\b/g, ' ');
  const numerosSueltos = notaSinRangosEdad.match(/\d+/g);

  const aliasesV0SinTalle = v0.aliases.filter((a) => !['1 a 3', '1-3', 't1', 'talle 1'].includes(a));
  const aliasesV1SinTalle = v1.aliases.filter((a) => !['3 a 6', '3-6', 't2', 'talle 2'].includes(a));

  const mencionaV0 =
    tieneRangoT1 || aliasesV0SinTalle.some((al) => new RegExp(`\\b${al}\\b`, 'i').test(notaNorm));
  const mencionaV1 =
    tieneRangoT2 || aliasesV1SinTalle.some((al) => new RegExp(`\\b${al}\\b`, 'i').test(notaNorm));

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

  // Parseo explícito por alias + cantidad (ej: "2 verde y 1 rosa", "t1: 2, t2: 1")
  const extraerCantidadVariante = (variante: VarianteCatalogo): number | null => {
    const aliasPattern = variante.aliases
      .map((a) => a.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+'))
      .join('|');
    const reDespues = new RegExp(`(?:${aliasPattern})\\s*(?::|=|-)?\\s*(\\d+)`, 'i');
    const reAntes = new RegExp(`(\\d+)\\s*(?:u\\.?|unid\\.?|de)?\\s*(?:${aliasPattern})\\b`, 'i');
    const m1 = notaNorm.match(reDespues);
    if (m1) return parseInt(m1[1], 10);
    const m2 = notaNorm.match(reAntes);
    if (m2) return parseInt(m2[1], 10);
    return null;
  };

  const c0 = extraerCantidadVariante(v0);
  const c1 = extraerCantidadVariante(v1);
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

const crearLineasEjemploLista006 = (): LineaPedido[] => {
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

const expandirCeldasCombinadas = (
  ws: SheetJSWorkSheet,
  XLSX: SheetJSModule
): CellPrimitive[][] => {
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
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (rows[r] && (rows[r][c] === null || rows[r][c] === '')) {
          rows[r][c] = topLeft;
        }
      }
    }
  }
  return rows;
};

const leerWorkbookSeguro = (
  buffer: ArrayBuffer,
  XLSX: SheetJSModule,
  fileName: string
): SheetJSWorkBook => {
  const esCsv = /\.csv$/i.test(fileName);
  if (esCsv) {
    const bytes = new Uint8Array(buffer);
    let texto = new TextDecoder('utf-8').decode(bytes);
    if (texto.includes('\uFFFD')) {
      texto = new TextDecoder('iso-8859-1').decode(bytes);
    }
    texto = limpiarTextoInvisible(texto);
    // Normalizar CSV argentino con punto y coma para que SheetJS lo parsee en columnas
    const lineasCsv = texto
      .split(/\r?\n/)
      .map((line) => line.split(';').map((cell) => limpiarTextoInvisible(cell.replace(/^"|"$/g, ''))));
    const ws = XLSX.utils.aoa_to_sheet(lineasCsv);
    return { SheetNames: ['CSV'], Sheets: { CSV: ws } };
  }
  return XLSX.read(buffer, { type: 'array' });
};

const parsearCantidadEstricta = (raw: CellPrimitive): { valid: boolean; qty: number; motivo?: string } => {
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

const parsearExcelLista006 = (
  buffer: ArrayBuffer,
  XLSX: SheetJSModule,
  fileName: string
): ResultadoParseoLista006 => {
  let wb: SheetJSWorkBook;
  try {
    wb = leerWorkbookSeguro(buffer, XLSX, fileName);
  } catch {
    return {
      ok: false,
      errorFatal: 'No se pudo leer el archivo. Verificá que sea un Excel (.xlsx, .xls) o CSV válido.',
      lineas: [],
      advertencias: []
    };
  }

  let esLista006 = false;
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = expandirCeldasCombinadas(ws, XLSX);
    for (const row of rows) {
      for (const cell of row) {
        if (cell === null) continue;
        const txt = normalizarTexto(String(cell)).replace(/["°º]/g, ' ');
        if (/lista\s*(?:n\s*)?0*6\b/.test(txt)) {
          esLista006 = true;
          break;
        }
      }
      if (esLista006) break;
    }
    if (esLista006) break;
  }

  if (!esLista006) {
    return {
      ok: false,
      errorFatal:
        'El archivo subido no corresponde a la plantilla oficial LISTA N° 006 de Mibi® (no se detectó la cabecera "LISTA N° 006").',
      lineas: [],
      advertencias: []
    };
  }

  const advertencias: AdvertenciaExcel[] = [];
  const lineasMap = new Map<string, LineaPedido>();

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = expandirCeldasCombinadas(ws, XLSX);

    let colCodeA = 0;
    let colCodeB = 1;
    let colPrecio = 3;
    let colQty = 4;
    let colNota = 6;
    let headerDetectado = false;

    for (let rIdx = 0; rIdx < rows.length; rIdx++) {
      const row = rows[rIdx];
      const rowTexts = row.map((c) => (c === null ? '' : normalizarTexto(String(c))));

      // Detectar encabezados en cualquier fila (sin límite de fila 30)
      const idxCantPed = rowTexts.findIndex((t) => t.includes('cantidades pedidas') || t.includes('cantidad pedida'));
      const idxCantFallback = rowTexts.findIndex((t) => t === 'cantidad' || t === 'cant');
      const idxTalleColor = rowTexts.findIndex(
        (t) => t.includes('talles/colores') || t.includes('talles / colores') || t.includes('talles/ colores')
      );
      const idxDetalleFallback = rowTexts.findIndex((t) => t.includes('detalle') && !t.includes('producto'));
      const idxCodigos = rowTexts.findIndex((t) => t === 'codigos' || t === 'codigo');
      const idxPrecio = rowTexts.findIndex((t) => t === 'precio' || t.includes('precio neto'));

      if (idxCantPed !== -1 || (idxCodigos !== -1 && (idxCantFallback !== -1 || idxPrecio !== -1))) {
        headerDetectado = true;
        if (idxCodigos !== -1) {
          colCodeA = idxCodigos;
          colCodeB = idxCodigos + 1;
        }
        if (idxPrecio !== -1) colPrecio = idxPrecio;
        if (idxCantPed !== -1) colQty = idxCantPed;
        else if (idxCantFallback !== -1) colQty = idxCantFallback;

        if (idxTalleColor !== -1) colNota = idxTalleColor;
        else if (idxDetalleFallback !== -1) colNota = idxDetalleFallback;
        continue;
      }

      // Buscar códigos de 4 dígitos en columnas de código o en primeras 2 columnas
      const rawA = row[colCodeA] !== null && row[colCodeA] !== undefined ? String(row[colCodeA]) : '';
      const rawB =
        colCodeB < row.length && row[colCodeB] !== null && row[colCodeB] !== undefined
          ? String(row[colCodeB])
          : '';
      const matchA = limpiarTextoInvisible(rawA).match(/\b(\d{4})\b/);
      const matchB = limpiarTextoInvisible(rawB).match(/\b(\d{4})\b/);

      const codigosEnFila: string[] = [];
      if (matchA) codigosEnFila.push(matchA[1]);
      if (matchB && matchB[1] !== matchA?.[1]) codigosEnFila.push(matchB[1]);

      if (codigosEnFila.length === 0) continue;

      const qtyCell = headerDetectado ? row[colQty] ?? null : row[4] ?? null;
      const precioCell = headerDetectado ? row[colPrecio] ?? null : row[3] ?? null;
      const notaCell = headerDetectado ? row[colNota] ?? null : row[6] ?? null;

      // Regla oficial: ignorar fila duplicada 6545 sin precio ni cantidad
      if (
        codigosEnFila.includes('6545') &&
        (precioCell === null || String(precioCell).trim() === '') &&
        (qtyCell === null || String(qtyCell).trim() === '' || Number(qtyCell) === 0)
      ) {
        continue;
      }

      const qtyRes = parsearCantidadEstricta(qtyCell);
      if (!qtyRes.valid) {
        advertencias.push({
          hoja: sheetName,
          fila: rIdx + 1,
          codigo: codigosEnFila.join('/'),
          mensaje: qtyRes.motivo ?? 'Cantidad inválida'
        });
        continue;
      }

      if (qtyRes.qty === 0) continue;

      // Buscar en catálogo
      let producto: ProductoCatalogo | undefined;
      for (const cod of codigosEnFila) {
        const encontrado = CATALOGO_BY_CODIGO.get(cod);
        if (encontrado) {
          producto = encontrado;
          break;
        }
      }

      if (!producto) {
        advertencias.push({
          hoja: sheetName,
          fila: rIdx + 1,
          codigo: codigosEnFila.join('/'),
          mensaje: `Código desconocido (${codigosEnFila.join('/')}) con cantidad ${qtyRes.qty}`
        });
        continue;
      }

      const notaStr = notaCell !== null && notaCell !== undefined ? String(notaCell) : '';
      const linea = resolverLinea(producto, qtyRes.qty, notaStr);
      const existente = lineasMap.get(producto.idGrupo);
      if (existente) {
        const suma = existente.cantidadTotal + linea.cantidadTotal;
        const notaCombinada = [existente.notaCliente, linea.notaCliente].filter(Boolean).join(' · ');
        lineasMap.set(producto.idGrupo, resolverLinea(producto, suma, notaCombinada));
      } else {
        lineasMap.set(producto.idGrupo, linea);
      }
    }
  }

  return {
    ok: true,
    lineas: Array.from(lineasMap.values()),
    advertencias
  };
};

// Parseo de números argentinos ("1.004.793,00", "690,573,06", "$50.000") evitando confundir fechas
const parsearMontoArgentino = (raw: CellPrimitive): number => {
  if (raw === null || raw === undefined) return 0;
  if (typeof raw === 'number') return Number.isFinite(raw) ? round2(raw) : 0;
  const s = limpiarTextoInvisible(String(raw));
  if (!s || /^#VALUE!$/i.test(s)) return 0;
  // Si es fecha pura (dd/mm o dd/mm/yyyy), no es un monto
  if (/^\d{1,2}\/\d{1,2}(?:\/\d{2,4})?$/.test(s)) return 0;

  // Extraer tokens tipo $50.000 o números con separadores (tolerando notas como "882,005,30 c Rocio")
  const sinSimboloRaw = s.replace(/\$/g, '').trim();
  const tokenNum = sinSimboloRaw.match(/^(-?\d[\d.,]*)/)?.[1] ?? '';
  const sinSimbolo = tokenNum || sinSimboloRaw;
  if (/^\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?$/.test(sinSimbolo)) {
    return round2(parseFloat(sinSimbolo.replace(/\./g, '').replace(',', '.')));
  }
  // Typo tipo "690,573,06" o "1,004,793,00" -> última coma es decimal
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

const parsearFechaExcel = (raw: CellPrimitive, fallbackISO: string): string => {
  if (raw === null || raw === undefined) return fallbackISO;
  if (typeof raw === 'number' && Number.isFinite(raw) && raw > 25000 && raw < 70000) {
    // Número serial de Excel (epoch 1899-12-30)
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
    const day = dmy[1].padStart(2, '0');
    const month = dmy[2].padStart(2, '0');
    let year = dmy[3] ? (dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3]) : String(new Date().getFullYear());
    if (year === '2006') year = '2026';
    return `${year}-${month}-${day}`;
  }
  return fallbackISO;
};

const importarPlanillaGeneral = (
  buffer: ArrayBuffer,
  XLSX: SheetJSModule,
  fileName: string,
  existentes: ReadonlyArray<Pedido>
): { ok: boolean; nuevos: Pedido[]; duplicadosOmitidos: number; error?: string } => {
  let wb: SheetJSWorkBook;
  try {
    wb = leerWorkbookSeguro(buffer, XLSX, fileName);
  } catch {
    return { ok: false, nuevos: [], duplicadosOmitidos: 0, error: 'No se pudo leer la planilla general.' };
  }

  const sheetName =
    wb.SheetNames.find((n) => normalizarTexto(n).includes('pedidos mayorista')) ?? wb.SheetNames[0];
  const ws = sheetName ? wb.Sheets[sheetName] : undefined;
  if (!ws) {
    return { ok: false, nuevos: [], duplicadosOmitidos: 0, error: 'La planilla no tiene hojas visibles.' };
  }

  const rows = expandirCeldasCombinadas(ws, XLSX);
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

  const headers = rows[headerIdx].map((h) => (h === null ? '' : normalizarTexto(String(h))));
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

  const nuevos: Pedido[] = [];
  let duplicadosOmitidos = 0;
  const hoy = getTodayISO();

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const numRaw =
      cNum !== -1 && r[cNum] !== null && r[cNum] !== undefined ? limpiarTextoInvisible(String(r[cNum])) : '';
    const cliRaw =
      cCli !== -1 && r[cCli] !== null && r[cCli] !== undefined ? limpiarTextoInvisible(String(r[cCli])) : '';
    if (!cliRaw || /^total\b/i.test(cliRaw) || /^total\b/i.test(numRaw)) continue;

    const numeroPedido = numRaw ? numRaw.replace(/^#/, '').padStart(4, '0') : String(existentes.length + nuevos.length + 1).padStart(4, '0');
    const clave = `${numeroPedido.replace(/^0+/, '')}|${normalizarTexto(cliRaw)}`;
    if (clavesExistentes.has(clave)) {
      duplicadosOmitidos += 1;
      continue;
    }
    clavesExistentes.add(clave);

    const totalFacturado = parsearMontoArgentino(cTot !== -1 ? r[cTot] ?? 0 : 0);
    const aboCell = cAbo !== -1 ? r[cAbo] ?? null : null;
    const aboVacio = aboCell === null || limpiarTextoInvisible(String(aboCell)) === '';
    let montoAbonado = parsearMontoArgentino(aboCell);
    if (montoAbonado > totalFacturado && totalFacturado > 0) montoAbonado = totalFacturado;

    const salCell = cSal !== -1 ? r[cSal] ?? null : null;
    const salStr = salCell !== null && salCell !== undefined ? limpiarTextoInvisible(String(salCell)) : '';
    const saldoSheet = salStr !== '' && !/^#VALUE!$/i.test(salStr) ? parsearMontoArgentino(salCell) : null;

    const estPagoRaw = cEstPago !== -1 && r[cEstPago] !== null && r[cEstPago] !== undefined ? limpiarTextoInvisible(String(r[cEstPago])) : '';
    const estPagoStr = normalizarTexto(estPagoRaw);
    const cheqRaw = cCheq !== -1 && r[cCheq] !== null && r[cCheq] !== undefined ? limpiarTextoInvisible(String(r[cCheq])) : '';
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

    const wspRaw = cWsp !== -1 && r[cWsp] !== null && r[cWsp] !== undefined ? limpiarTextoInvisible(String(r[cWsp])) : '';
    const vendRaw =
      cVend !== -1 && r[cVend] !== null && r[cVend] !== undefined && limpiarTextoInvisible(String(r[cVend]))
        ? limpiarTextoInvisible(String(r[cVend]))
        : !/^\d+$/.test(wspRaw.replace(/[\s\-+]/g, '')) && wspRaw
        ? wspRaw
        : 'Administración';

    const armRaw = normalizarTexto(String(cArm !== -1 ? r[cArm] ?? '' : ''));
    const estadoArmado: EstadoArmado = armRaw.includes('listo')
      ? 'Listo'
      : armRaw.includes('proceso') || (armRaw.includes('armado') && armRaw.includes('en'))
      ? 'En armado'
      : 'Por armar';

    const envRaw = normalizarTexto(String(cEnv !== -1 ? r[cEnv] ?? '' : ''));
    const estadoEnvio: EstadoEnvio = envRaw.includes('entreg')
      ? 'Entregado'
      : envRaw.includes('despach')
      ? 'Despachado'
      : 'Pendiente';

    const fechaPedido = parsearFechaExcel(cFec !== -1 ? r[cFec] ?? null : null, hoy);
    const despachoPrevisto = parsearFechaExcel(cDesp !== -1 ? r[cDesp] ?? null : null, '');
    const fechaDespacho = parsearFechaExcel(cFecDesp !== -1 ? r[cFecDesp] ?? null : null, '');
    const fechaEntrega = parsearFechaExcel(cFecEnt !== -1 ? r[cFecEnt] ?? null : null, '');

    const ped = recalcularPedido({
      id: `imp_${Date.now()}_${i}`,
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
      detalleCheque: esCheque ? (cheqRaw || estPagoRaw) : '',
      estadoArmado,
      estadoEnvio,
      transporte: cTrans !== -1 && r[cTrans] ? limpiarTextoInvisible(String(r[cTrans])) : '',
      nroSeguimiento: cGuia !== -1 && r[cGuia] ? limpiarTextoInvisible(String(r[cGuia])) : '',
      observaciones: cObs !== -1 && r[cObs] ? limpiarTextoInvisible(String(r[cObs])) : '',
      createdAt: Date.now() - i,
      updatedAt: Date.now(),
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

// ============================================================================
// 3 EXPORTADORES EXCEL .XLSX REALES (CELDAS NUMÉRICAS t:'n' — CERO #VALUE!)
// ============================================================================

const aplicarFormatoMoneda = (ws: SheetJSWorkSheet, cols: ReadonlyArray<number>, XLSX: SheetJSModule): void => {
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

// 1. Hoja para Depósito (individual) — bloqueo estricto si faltantesVariante > 0
const exportarHojaDeposito = (pedido: Pedido, XLSX: SheetJSModule | null): ResultadoHojaDeposito => {
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
  if (!XLSX) {
    return { success: false, errors: ['Preparando motor Excel, intentá en un segundo.'] };
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
    ['Transporte', pedido.transporte || 'A definir', 'Despacho previsto', pedido.despachoPrevisto, 'WhatsApp', pedido.whatsapp || '—'],
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
  aplicarFormatoMoneda(ws, [4, 5], XLSX);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `Pedido_${pedido.numeroPedido}`);
  XLSX.writeFile(wb, `Mibi_Deposito_Pedido_${pedido.numeroPedido}_${getTodayISO()}.xlsx`);
  return { success: true, errors: [] };
};

const isoToExcelSerial = (iso: string): number | null => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return null;
  const [y, m, d] = iso.split('-').map(Number);
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

const descargarBlob = (blob: Blob, filename: string): void => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.setTimeout(() => URL.revokeObjectURL(url), 1500);
};

// 2. Tablero Maestro PEDIDOS MAYORISTA MIBI.xlsx (preserva plantilla oficial con 2 hojas, estilos, colores, fórmulas y fechas Excel)
const exportarTableroMaestro = async (
  pedidos: ReadonlyArray<Pedido>,
  resumen: ResumenFinanciero,
  XLSX: SheetJSModule,
  sufijoArchivo = ''
): Promise<void> => {
  const sufijo = sufijoArchivo ? `_${sufijoArchivo}` : '';
  const nombreArchivo = `PEDIDOS_MAYORISTA_MIBI${sufijo}_${getTodayISO()}.xlsx`;

  const pedidosOrdenados = [...pedidos].sort((a, b) => {
    const na = parseInt(a.numeroPedido.replace(/\D/g, ''), 10);
    const nb = parseInt(b.numeroPedido.replace(/\D/g, ''), 10);
    if (Number.isFinite(na) && Number.isFinite(nb) && na !== nb) return na - nb;
    return a.numeroPedido.localeCompare(b.numeroPedido);
  });

  // Intentar exportar sobre la plantilla oficial exacta (PEDIDOS MAYORISTA MIBI + Vista Compacta)
  try {
    const resp = typeof window !== 'undefined' ? await fetch('/plantilla_pedidos_mibi.xlsx') : null;
    if (resp && resp.ok) {
      const tplBuf = await resp.arrayBuffer();
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
          const plain = m[1].replace(/<[^>]+>/g, '');
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
                  p.detalleCheque || (p.fechaAcreditacionCheque ? `Cheque (${formatShortDate(p.fechaAcreditacionCheque)})` : 'Cheque')
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

        // Actualizar Sheet 1 (filas 11..299+)
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
        const s1End = s1Xml.indexOf('<row r="300"');
        if (s1Start !== -1 && s1End !== -1) {
          s1Xml = s1Xml.slice(0, s1Start) + rowsS1.join('') + s1Xml.slice(s1End);
        }
        zip.file('xl/worksheets/sheet1.xml', s1Xml);

        // Actualizar Sheet 2 (Vista Compacta, filas 4..303+)
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
        const s2End = s2Xml.indexOf('<row r="304"');
        if (s2Start !== -1 && s2End !== -1) {
          s2Xml = s2Xml.slice(0, s2Start) + rowsS2.join('') + s2Xml.slice(s2End);
        }
        zip.file('xl/worksheets/sheet2.xml', s2Xml);

        // Actualizar sharedStrings.xml
        const newSsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${siList.length}" uniqueCount="${siList.length}">${siList.join('')}</sst>`;
        zip.file('xl/sharedStrings.xml', newSsXml);

        const outBytes = await zip.generateAsync({ type: 'uint8array' });
        const outBlob = new Blob([outBytes.buffer as ArrayBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        descargarBlob(outBlob, nombreArchivo);
        return;
      }
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
  aplicarFormatoMoneda(ws, [6, 7, 8], XLSX);

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
  XLSX.writeFile(wb, nombreArchivo);
};

// 3. Consolidado para Depósito (Picking Unificado por SKU + Desglose por Cliente)
// Regla estricta: excluye automáticamente pedidos con talle/color sin definir o sin líneas
const exportarPickingConsolidado = (
  seleccionados: ReadonlyArray<Pedido>,
  XLSX: SheetJSModule | null
): ResultadoPickingConsolidado => {
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

  if (!XLSX) {
    return {
      success: false,
      incluidos: 0,
      excluidosTalle,
      excluidosSinLineas,
      totalUnidades: 0,
      error: 'Preparando motor Excel, intentá en un segundo.'
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
  XLSX.writeFile(wb, `Mibi_Picking_Consolidado_${validos.length}ped_${getTodayISO()}.xlsx`);

  return {
    success: true,
    incluidos: validos.length,
    excluidosTalle,
    excluidosSinLineas,
    totalUnidades
  };
};

// ============================================================================
// VALIDACIÓN ZOD, PERSISTENCIA HÍBRIDA (NUBE FIREBASE + LOCALSTORAGE FALLBACK)
// ============================================================================

const DesgloseSchema = z.object({
  codigo: z.string().min(1),
  etiqueta: z.string(),
  cantidad: z.number().int().min(0)
});

const LineaPedidoSchema = z.object({
  idLinea: z.string().min(1),
  idGrupo: z.string().min(1),
  nombreProducto: z.string().min(1),
  rubro: z.string(),
  precioUnitarioNeto: z.number().min(0),
  cantidadTotal: z.number().int().min(1),
  notaCliente: z.string(),
  estado: z.enum(['OK', 'REQUIERE_REVISION']),
  requiereRevision: z.boolean(),
  desglose: z.array(DesgloseSchema)
});

const PedidoSchema = z.object({
  id: z.string().min(1),
  numeroPedido: z.string().min(1),
  cliente: z.string().min(1),
  vendedor: z.string(),
  whatsapp: z.string(),
  fechaPedido: z.string(),
  despachoPrevisto: z.string(),
  lineas: z.array(LineaPedidoSchema),
  descuentoPorc: z.number().min(0).max(100),
  medioPago: z.enum(['Transferencia / Efectivo', 'Cheque', 'Sin pagar']),
  montoAbonado: z.number().min(0),
  montoCheque: z.number().min(0),
  chequeAcreditado: z.boolean(),
  estadoArmado: z.enum(['Por armar', 'En armado', 'Listo', 'Falta talle o color']),
  estadoEnvio: z.enum(['Pendiente', 'Despachado', 'Entregado']),
  totalFacturado: z.number().min(0)
});

const validarPedidoSchema = (pedido: Pedido): { success: boolean } => PedidoSchema.safeParse(pedido).success ? { success: true } : { success: false };

const initFirebaseSafely = (): { app: FirebaseApp | null; db: Firestore | null } => {
  try {
    const env = typeof process !== 'undefined' ? process.env : undefined;
    const apiKey = env?.NEXT_PUBLIC_FIREBASE_API_KEY;
    const projectId = env?.NEXT_PUBLIC_FIREBASE_PROJECT_ID;
    const authDomain = env?.NEXT_PUBLIC_FIREBASE_AUTH_DOMAIN;
    if (!apiKey || !projectId) {
      return { app: null, db: null };
    }
    const existing = getApps();
    const app =
      existing.length > 0
        ? existing[0]
        : initializeApp({
            apiKey,
            authDomain: authDomain || `${projectId}.firebaseapp.com`,
            projectId
          });
    return { app, db: getFirestore(app) };
  } catch {
    return { app: null, db: null };
  }
};

const { app: firebaseApp, db: cloudDb } = initFirebaseSafely();
const CLOUD_APP_ID = 'mibi_erp_prod';

const getPublicOrdersCollection = (db: Firestore) =>
  collection(db, 'artifacts', CLOUD_APP_ID, 'public', 'data', 'orders');

const getPublicOrderDoc = (db: Firestore, orderId: string): DocumentReference =>
  doc(db, 'artifacts', CLOUD_APP_ID, 'public', 'data', 'orders', orderId);

const getTeamConfigDoc = (db: Firestore): DocumentReference =>
  doc(db, 'artifacts', CLOUD_APP_ID, 'public', 'data', 'config', 'team');

const getUserSessionDoc = (db: Firestore, uid: string): DocumentReference =>
  doc(db, 'artifacts', CLOUD_APP_ID, 'users', uid, 'session', 'current');

const serializePedidoForCloud = (p: Pedido): Record<string, unknown> => ({
  id: p.id,
  numeroPedido: p.numeroPedido,
  cliente: p.cliente,
  vendedor: p.vendedor,
  whatsapp: p.whatsapp,
  fechaPedido: p.fechaPedido,
  despachoPrevisto: p.despachoPrevisto,
  lineas: p.lineas,
  descuentoPorc: p.descuentoPorc,
  medioPago: p.medioPago,
  montoAbonado: p.montoAbonado,
  montoCheque: p.montoCheque,
  chequeAcreditado: p.chequeAcreditado,
  fechaRecepcionCheque: p.fechaRecepcionCheque,
  fechaAcreditacionCheque: p.fechaAcreditacionCheque,
  detalleCheque: p.detalleCheque,
  estadoArmado: p.estadoArmado,
  estadoEnvio: p.estadoEnvio,
  fechaDespacho: p.fechaDespacho ?? '',
  fechaEntrega: p.fechaEntrega ?? '',
  transporte: p.transporte,
  nroSeguimiento: p.nroSeguimiento,
  observaciones: p.observaciones,
  createdAt: p.createdAt,
  updatedAt: p.updatedAt,
  historial: p.historial,
  detalleManualImportado: p.detalleManualImportado ?? '',
  totalFacturadoOverride: p.totalFacturadoOverride ?? null
});

const deserializePedidoFromCloud = (docId: string, raw: Record<string, unknown>): Pedido | null => {
  const cliente = readString(raw, 'cliente');
  if (!cliente) return null;

  const medioRaw = readString(raw, 'medioPago', 'Sin pagar');
  const medioPago: MedioPago =
    medioRaw === 'Transferencia / Efectivo' || medioRaw === 'Cheque' || medioRaw === 'Sin pagar'
      ? medioRaw
      : 'Sin pagar';

  const armadoRaw = readString(raw, 'estadoArmado', 'Por armar');
  const estadoArmado: EstadoArmado =
    armadoRaw === 'Por armar' ||
    armadoRaw === 'En armado' ||
    armadoRaw === 'Listo' ||
    armadoRaw === 'Falta talle o color'
      ? armadoRaw
      : 'Por armar';

  const envioRaw = readString(raw, 'estadoEnvio', 'Pendiente');
  const estadoEnvio: EstadoEnvio =
    envioRaw === 'Pendiente' || envioRaw === 'Despachado' || envioRaw === 'Entregado' ? envioRaw : 'Pendiente';

  const rawLineas = Array.isArray(raw['lineas']) ? raw['lineas'] : [];
  const lineas: LineaPedido[] = rawLineas
    .filter((item): item is Record<string, unknown> => typeof item === 'object' && item !== null)
    .map((l) => {
      const estadoL: EstadoLinea = readString(l, 'estado') === 'OK' ? 'OK' : 'REQUIERE_REVISION';
      const desgloseRaw = Array.isArray(l['desglose']) ? l['desglose'] : [];
      const desglose: DesgloseVariante[] = desgloseRaw
        .filter((d): d is Record<string, unknown> => typeof d === 'object' && d !== null)
        .map((d) => ({
          codigo: readString(d, 'codigo'),
          etiqueta: readString(d, 'etiqueta'),
          cantidad: readNumber(d, 'cantidad', 0)
        }));
      return {
        idLinea: readString(l, 'idLinea'),
        idGrupo: readString(l, 'idGrupo'),
        nombreProducto: readString(l, 'nombreProducto'),
        rubro: readString(l, 'rubro'),
        precioUnitarioNeto: readNumber(l, 'precioUnitarioNeto', 0),
        cantidadTotal: readNumber(l, 'cantidadTotal', 1),
        notaCliente: readString(l, 'notaCliente'),
        estado: estadoL,
        requiereRevision: readBoolean(l, 'requiereRevision', estadoL !== 'OK'),
        motivoRevision: readString(l, 'motivoRevision') || undefined,
        desglose
      };
    });

  const rawHist = Array.isArray(raw['historial']) ? raw['historial'] : [];
  const historial: HistorialEvento[] = rawHist
    .filter((h): h is Record<string, unknown> => typeof h === 'object' && h !== null)
    .map((h) => ({
      fechaHora: readString(h, 'fechaHora'),
      usuario: readString(h, 'usuario'),
      rol: readString(h, 'rol') === 'Administración' ? 'Administración' : 'Ventas / Depósito',
      accion: readString(h, 'accion')
    }));

  const overrideVal = raw['totalFacturadoOverride'];
  const ped = recalcularPedido({
    id: docId,
    numeroPedido: readString(raw, 'numeroPedido', docId),
    cliente,
    vendedor: readString(raw, 'vendedor', 'Administración'),
    whatsapp: readString(raw, 'whatsapp'),
    fechaPedido: readString(raw, 'fechaPedido', getTodayISO()),
    despachoPrevisto: readString(raw, 'despachoPrevisto', addDaysISO(getTodayISO(), 5)),
    lineas,
    descuentoPorc: readNumber(raw, 'descuentoPorc', 0),
    medioPago,
    montoAbonado: readNumber(raw, 'montoAbonado', 0),
    montoCheque: readNumber(raw, 'montoCheque', 0),
    chequeAcreditado: readBoolean(raw, 'chequeAcreditado', false),
    fechaRecepcionCheque: readString(raw, 'fechaRecepcionCheque'),
    fechaAcreditacionCheque: readString(raw, 'fechaAcreditacionCheque'),
    detalleCheque: readString(raw, 'detalleCheque'),
    estadoArmado,
    estadoEnvio,
    fechaDespacho: readString(raw, 'fechaDespacho') || undefined,
    fechaEntrega: readString(raw, 'fechaEntrega') || undefined,
    transporte: readString(raw, 'transporte'),
    nroSeguimiento: readString(raw, 'nroSeguimiento'),
    observaciones: readString(raw, 'observaciones'),
    createdAt: readNumber(raw, 'createdAt', Date.now()),
    updatedAt: readNumber(raw, 'updatedAt', Date.now()),
    historial,
    detalleManualImportado: readString(raw, 'detalleManualImportado') || undefined,
    totalFacturadoOverride: typeof overrideVal === 'number' && Number.isFinite(overrideVal) ? overrideVal : undefined
  });

  return validarPedidoSchema(ped).success ? ped : null;
};

const saveOrderInCloud = async (db: Firestore, pedido: Pedido): Promise<void> => {
  await setDoc(getPublicOrderDoc(db, pedido.id), serializePedidoForCloud(pedido));
};

const patchOrderInCloud = async (
  db: Firestore,
  orderId: string,
  patch: Partial<PedidoBase>,
  historial: HistorialEvento[],
  updatedAt: number
): Promise<void> => {
  await updateDoc(getPublicOrderDoc(db, orderId), {
    ...patch,
    historial,
    updatedAt
  });
};

const clearAllOrdersInCloud = async (db: Firestore, orderIds: ReadonlyArray<string>): Promise<void> => {
  const CHUNK_SIZE = 400;
  for (let i = 0; i < orderIds.length; i += CHUNK_SIZE) {
    const batch = writeBatch(db);
    for (const id of orderIds.slice(i, i + CHUNK_SIZE)) {
      batch.delete(getPublicOrderDoc(db, id));
    }
    await batch.commit();
  }
};

const saveTeamConfigToCloud = async (db: Firestore, config: ConfigEquipo): Promise<void> => {
  await setDoc(getTeamConfigDoc(db), config);
};

const saveUserSessionToCloud = async (db: Firestore, uid: string, sesion: SesionUsuario): Promise<void> => {
  await setDoc(getUserSessionDoc(db, uid), sesion);
};

const clearUserSessionInCloud = async (db: Firestore, uid: string): Promise<void> => {
  await setDoc(getUserSessionDoc(db, uid), { nombre: '', rol: 'Ventas / Depósito', configVersion: 0 });
};

// Persistencia híbrida localStorage (respaldo offline + filtros persistentes + modo local sin Firebase)
const LS_KEYS = {
  pedidos: 'mibi_erp_pedidos_v1',
  filtros: 'mibi_erp_filtros_v1',
  config: 'mibi_erp_config_v1',
  sesion: 'mibi_erp_sesion_v1'
} as const;

const guardarJsonLocal = (key: string, value: unknown): void => {
  try {
    if (typeof window !== 'undefined' && window.localStorage) {
      window.localStorage.setItem(key, JSON.stringify(value));
    }
  } catch {
    // Modo privado o cuota llena: no interrumpe operación
  }
};

const leerJsonLocal = (key: string): unknown => {
  try {
    if (typeof window !== 'undefined' && window.localStorage) {
      const raw = window.localStorage.getItem(key);
      return raw ? JSON.parse(raw) : null;
    }
  } catch {
    return null;
  }
  return null;
};

const cargarPedidosLocales = (): Pedido[] => {
  const raw = leerJsonLocal(LS_KEYS.pedidos);
  if (!Array.isArray(raw)) return [];
  const validos: Pedido[] = [];
  for (const item of raw) {
    if (typeof item !== 'object' || item === null) continue;
    const rec = recalcularPedido(item as PedidoBase);
    if (validarPedidoSchema(rec).success) validos.push(rec);
  }
  return validos;
};

const cargarFiltrosLocales = (): { busqueda: string; filtro: FiltroPedidos } => {
  const raw = leerJsonLocal(LS_KEYS.filtros);
  if (typeof raw !== 'object' || raw === null) return { busqueda: '', filtro: 'todos' };
  const obj = raw as Record<string, unknown>;
  const busqueda = typeof obj['busqueda'] === 'string' ? obj['busqueda'] : '';
  const f = obj['filtro'];
  const filtro: FiltroPedidos =
    f === 'todos' || f === 'falta_talle' || f === 'cheques' || f === 'falta_pago' || f === 'listos'
      ? f
      : 'todos';
  return { busqueda, filtro };
};

// ============================================================================
// MENSAJES AUTOMÁTICOS WHATSAPP Y SUITE QA EN VIVO (27 PRUEBAS)
// ============================================================================

const generarMensajeWhatsApp = (
  cliente: string,
  numeroPedido: string,
  lineas: ReadonlyArray<LineaPedido>
): string => {
  const pendientes = lineas.filter((l) => l.requiereRevision || l.estado !== 'OK');
  if (pendientes.length === 0) return '';
  const lista = pendientes
    .map((l) => {
      const prod = CATALOGO_BY_GRUPO.get(l.idGrupo);
      const opciones = prod ? prod.variantes.map((v) => `${v.etiqueta} (${v.codigo})`).join(' / ') : '';
      return `• ${l.nombreProducto} — ${l.cantidadTotal} u. (elegir entre: ${opciones})`;
    })
    .join('\n');

  const saludo = cliente ? `¡Hola ${cliente}!` : '¡Hola!';
  const ref = numeroPedido ? ` por tu pedido #${numeroPedido}` : ' por tu pedido';
  return `${saludo} Te escribimos de Mibi®${ref}.\n\nPara pasarlo a depósito sin errores, nos falta confirmar talle o color en:\n\n${lista}\n\nApenas nos confirmes esas cantidades lo dejamos listo para armar. ¡Muchas gracias!`;
};

const generarMensajeDespachoWhatsApp = (pedido: Pedido): string => {
  const transporte = pedido.transporte.trim();
  const guia = pedido.nroSeguimiento.trim();
  if (!transporte || !guia) return '';
  return `¡Hola ${pedido.cliente}! Tu pedido #${pedido.numeroPedido} de Mibi® ya fue despachado por ${transporte}. Guía de seguimiento: ${guia}`;
};

const ejecutarSuiteQA = (XLSX: SheetJSModule | null): ResultadoQA[] => {
  const resultados: ResultadoQA[] = [];
  const check = (nombre: string, ok: boolean, detalle: string): void => {
    resultados.push({ nombre, ok, detalle });
  };

  const monster = CATALOGO_BY_CODIGO.get('6502')!;
  const shark = CATALOGO_BY_CODIGO.get('6504')!;
  const zapa = CATALOGO_BY_CODIGO.get('6576')!;
  const dino = CATALOGO_BY_CODIGO.get('6510')!;

  // 1-4. Redondeos IEEE-754 y aritmética contable
  check('Redondeo 0.1 + 0.2 === 0.30', round2(0.1 + 0.2) === 0.3, `Resultado: ${round2(0.1 + 0.2)}`);
  check('Redondeo 35.855 === 35.86', round2(35.855) === 35.86, `Resultado: ${round2(35.855)}`);
  check('Redondeo 1.005 === 1.01', round2(1.005) === 1.01, `Resultado: ${round2(1.005)}`);
  check('Catálogo completo (6500-6614)', CATALOGO_BY_CODIGO.size >= 62, `${CATALOGO_BY_CODIGO.size} códigos activos`);

  // 5-13. Reglas estrictas de desglose
  const lUnica = resolverLinea(dino, 3, '');
  check('Variante única (Antiparras Dino) → OK automático', lUnica.estado === 'OK' && lUnica.desglose[0].cantidad === 3, lUnica.estado);

  const lSimetrica = resolverLinea(shark, 6, '3 y 3');
  check('Nota simétrica "3 y 3" sobre 6 → OK (3 y 3)', lSimetrica.estado === 'OK' && lSimetrica.desglose[0].cantidad === 3 && lSimetrica.desglose[1].cantidad === 3, lSimetrica.estado);

  const lAsimetrica = resolverLinea(shark, 3, '2 y 1');
  check('Nota asimétrica "2 y 1" → REQUIERE_REVISION', lAsimetrica.requiereRevision, lAsimetrica.motivoRevision ?? '');

  const lRango = resolverLinea(shark, 3, 'de 1 a 3');
  check('Nota "de 1 a 3" sobre 3 → todo a T1 (6504)', lRango.estado === 'OK' && lRango.desglose[0].cantidad === 3, lRango.desglose[0].codigo);

  const lRangoMismatch = resolverLinea(shark, 4, 'de 1 a 3: 2');
  check('Nota "de 1 a 3: 2" sobre 4 → REQUIERE_REVISION', lRangoMismatch.requiereRevision, lRangoMismatch.motivoRevision ?? '');

  const lGris = resolverLinea(monster, 3, 'gris');
  check('Nota "gris" en Monster → REQUIERE_REVISION', lGris.requiereRevision, lGris.motivoRevision ?? '');

  const lSurtido = resolverLinea(monster, 4, 'surtido');
  check('Nota "surtido" → REQUIERE_REVISION', lSurtido.requiereRevision, lSurtido.motivoRevision ?? '');

  const lVacia = resolverLinea(monster, 3, '');
  check('Nota vacía en 2 variantes → REQUIERE_REVISION', lVacia.requiereRevision, lVacia.motivoRevision ?? '');

  const lImpar = resolverLinea(monster, 3, 'mitad y mitad');
  check('"mitad y mitad" impar (3) → REQUIERE_REVISION', lImpar.requiereRevision, lImpar.motivoRevision ?? '');

  // 14-18. Finanzas y tope de monto abonado
  const pedBase: PedidoBase = {
    id: 'qa_1',
    numeroPedido: '9990',
    cliente: 'Cliente QA',
    vendedor: 'Administración',
    whatsapp: '',
    fechaPedido: getTodayISO(),
    despachoPrevisto: addDaysISO(getTodayISO(), 5),
    lineas: crearLineasEjemploLista006(),
    descuentoPorc: 0,
    medioPago: 'Sin pagar',
    montoAbonado: 0,
    montoCheque: 0,
    chequeAcreditado: false,
    fechaRecepcionCheque: '',
    fechaAcreditacionCheque: '',
    detalleCheque: '',
    estadoArmado: 'Por armar',
    estadoEnvio: 'Pendiente',
    transporte: '',
    nroSeguimiento: '',
    observaciones: '',
    createdAt: Date.now(),
    updatedAt: Date.now(),
    historial: []
  };

  const pedDemo = recalcularPedido(pedBase);
  check(
    'Caso real Lista 006: $2.942.705 neto / $617.968,05 IVA / $3.560.673,05 total',
    pedDemo.subtotalNeto === 2942705 && pedDemo.iva21 === 617968.05 && pedDemo.totalFacturado === 3560673.05,
    `${formatARS(pedDemo.totalFacturado)}`
  );

  const pedTope = recalcularPedido({ ...pedBase, lineas: [lUnica], montoAbonado: 99999999 });
  check('Tope obligatorio: montoAbonado <= totalFacturado', pedTope.montoAbonado === pedTope.totalFacturado && pedTope.saldoPendiente === 0, `${formatARS(pedTope.montoAbonado)}`);

  const pedCheque = recalcularPedido({
    ...pedBase,
    lineas: [lUnica],
    medioPago: 'Cheque',
    montoCheque: pedTope.totalFacturado,
    chequeAcreditado: false,
    fechaRecepcionCheque: getTodayISO(),
    fechaAcreditacionCheque: addDaysISO(getTodayISO(), 30)
  });
  check('Cheque sin acreditar → estado Pendiente', pedCheque.estadoPago === 'Pendiente' && pedCheque.deudaSinRespaldo === 0, pedCheque.estadoPago);

  const pedChequeVencido = recalcularPedido({ ...pedCheque, fechaAcreditacionCheque: addDaysISO(getTodayISO(), -2) });
  check('Cheque vencido → alerta VENCIDO', getChequeAlertStatus(pedChequeVencido) === 'VENCIDO', getChequeAlertStatus(pedChequeVencido) ?? '');

  const pedChequeHoy = recalcularPedido({ ...pedCheque, fechaAcreditacionCheque: getTodayISO() });
  check(
    'Cheque con vencimiento hoy → 0 días (Al cobro)',
    getDiasParaAcreditacionCheque(pedChequeHoy) === 0 && getChequeAlertStatus(pedChequeHoy) === 'VENCE_HOY',
    `Días: ${getDiasParaAcreditacionCheque(pedChequeHoy)}`
  );

  // 19-24. Validaciones Parser (CSV, inválidos, negativos, invisibles, desconocidos)
  const qInvalida = parsearCantidadEstricta('3 docenas');
  check('Rechaza cantidad "3 docenas"', !qInvalida.valid, qInvalida.motivo ?? '');

  const qNegativa = parsearCantidadEstricta('-2');
  check('Rechaza cantidad negativa "-2"', !qNegativa.valid, qNegativa.motivo ?? '');

  const qDecimal = parsearCantidadEstricta('1,5');
  check('Rechaza cantidad decimal "1,5"', !qDecimal.valid, qDecimal.motivo ?? '');

  const limpioInv = limpiarTextoInvisible('\u200B6504\uFEFF');
  check('Limpia caracteres invisibles \\u200B/\\uFEFF', limpioInv === '6504', limpioInv);

  const lExplicita = resolverLinea(zapa, 4, '2 t1 y 2 t2');
  check('Desglose explícito "2 t1 y 2 t2" → OK', lExplicita.estado === 'OK' && lExplicita.desglose[0].cantidad === 2 && lExplicita.desglose[1].cantidad === 2, lExplicita.estado);

  const desc10 = recalcularPedido({ ...pedBase, lineas: [lUnica], descuentoPorc: 10 });
  const baseEsp = round2(round2(3 * 9505) * 0.9);
  const totEsp = round2(baseEsp * 1.21);
  check('Descuento 10% + IVA 21% exacto', desc10.baseImponible === baseEsp && desc10.totalFacturado === totEsp, formatARS(desc10.totalFacturado));

  // 25-27. Bloqueos de depósito, demo Lista 006 y WhatsApp
  check('Demo Lista 006: 1 artículo frenado (Monster)', pedDemo.faltantesVariante === 1 && pedDemo.estadoArmado === 'Falta talle o color', `Faltantes: ${pedDemo.faltantesVariante}`);

  const pedConPendiente = recalcularPedido({ ...pedBase, numeroPedido: '9991', lineas: [lVacia] });
  const hojaBloqueada = exportarHojaDeposito(pedConPendiente, XLSX);
  check('Hoja para Depósito bloqueada si falta talle/color', !hojaBloqueada.success, hojaBloqueada.errors[0] ?? '');

  const msgDespacho = generarMensajeDespachoWhatsApp({
    ...pedCheque,
    cliente: 'Playa Kids',
    numeroPedido: '0042',
    transporte: 'Andreani',
    nroSeguimiento: 'AND-9981'
  });
  check(
    'Mensaje WhatsApp de envío con transporte y guía',
    msgDespacho.includes('Playa Kids') && msgDespacho.includes('#0042') && msgDespacho.includes('Andreani') && msgDespacho.includes('AND-9981'),
    'Texto listo para copiar'
  );

  // 28. Picking consolidado excluye/frena pedidos sin talle
  const pickBloqueado = exportarPickingConsolidado([pedConPendiente], XLSX);
  check(
    'Picking consolidado frena/excluye pedidos sin talle',
    !pickBloqueado.success && pickBloqueado.excluidosTalle.includes('#9991'),
    pickBloqueado.error ?? 'Bloqueado'
  );

  // 29. Monto argentino con fechas en observaciones
  const montoFecha = parsearMontoArgentino('10/11');
  check('No confunde fecha "10/11" con monto', montoFecha === 0, `Resultado: ${montoFecha}`);

  return resultados;
};

// ============================================================================
// COMPONENTE UI AUXILIAR (STATUS PILL)
// ============================================================================

function StatusPill({
  tone = 'neutral',
  children
}: {
  tone?: ToneStatus;
  children: React.ReactNode;
}): React.JSX.Element {
  const styles: Record<ToneStatus, { box: string; dot: string }> = {
    green: {
      box: 'bg-emerald-50 text-emerald-700 border-emerald-200/60',
      dot: 'bg-emerald-500'
    },
    amber: {
      box: 'bg-amber-50 text-amber-700 border-amber-200/60',
      dot: 'bg-amber-500'
    },
    red: {
      box: 'bg-red-50 text-red-700 border-red-200/60',
      dot: 'bg-red-500'
    },
    neutral: {
      box: 'bg-zinc-100 text-zinc-700 border-zinc-200/70',
      dot: 'bg-zinc-400'
    },
    ink: {
      box: 'bg-zinc-900/[0.06] text-zinc-900 border-zinc-300/80',
      dot: 'bg-zinc-900'
    }
  };
  const s = styles[tone];
  return (
    <span
      className={`inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-medium border whitespace-nowrap ${s.box}`}
    >
      <span className={`w-1.5 h-1.5 rounded-full ${s.dot}`} />
      {children}
    </span>
  );
}

// ============================================================================
// APLICACIÓN PRINCIPAL — MIBI® ERP MAYORISTA
// ============================================================================

const crearFormInicial = (): FormNuevoPedido => {
  const hoy = getTodayISO();
  return {
    cliente: '',
    vendedor: 'Juancho',
    whatsapp: '',
    fechaPedido: hoy,
    despachoPrevisto: addDaysISO(hoy, 5),
    observaciones: '',
    descuentoPorc: 0,
    pagoTipo: 'Sin pagar',
    montoAbonadoInput: '',
    montoChequeInput: '',
    fechaRecepcionCheque: hoy,
    fechaAcreditacionCheque: addDaysISO(hoy, 30),
    detalleCheque: ''
  };
};

export default function App(): React.JSX.Element {
  // 1. Estados base de montaje, SheetJS y nube
  const [mounted, setMounted] = useState<boolean>(false);
  const [XLSX, setXLSX] = useState<SheetJSModule | null>(null);
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [cargandoNube, setCargandoNube] = useState<boolean>(Boolean(cloudDb));
  const [errorNube, setErrorNube] = useState<boolean>(false);
  const [pedidosInvalidosNube, setPedidosInvalidosNube] = useState<number>(0);

  // 2. Estados de equipo, claves y sesión
  const [configEquipo, setConfigEquipo] = useState<ConfigEquipo | null>(null);
  const [sesion, setSesion] = useState<SesionUsuario | null>(null);
  const [nombreIngreso, setNombreIngreso] = useState<string>('');
  const [pinIngreso, setPinIngreso] = useState<string>('');
  const [pinAdminNuevo, setPinAdminNuevo] = useState<string>('');
  const [pinVentasNuevo, setPinVentasNuevo] = useState<string>('');
  const [errorIngreso, setErrorIngreso] = useState<string>('');
  const [guardandoAcceso, setGuardandoAcceso] = useState<boolean>(false);
  const [modalEquipoAbierto, setModalEquipoAbierto] = useState<boolean>(false);
  const [pinAdminModal, setPinAdminModal] = useState<string>('');
  const [pinVentasModal, setPinVentasModal] = useState<string>('');
  const [resultadosQA, setResultadosQA] = useState<ResultadoQA[] | null>(null);

  // 3. Estados de pedidos, navegación, filtros y drawer (declarados ANTES de cualquier useEffect)
  const [pedidos, setPedidos] = useState<Pedido[]>([]);
  const [vista, setVista] = useState<VistaActiva>('pedidos');
  const [busqueda, setBusqueda] = useState<string>('');
  const [filtro, setFiltro] = useState<FiltroPedidos>('todos');
  const [drawerPedidoId, setDrawerPedidoId] = useState<string | null>(null);
  const [drawerDraft, setDrawerDraft] = useState<DrawerDraft | null>(null);
  const [mostrarHojaImprimible, setMostrarHojaImprimible] = useState<boolean>(false);
  const [toast, setToast] = useState<ToastState | null>(null);
  const [copiadoId, setCopiadoId] = useState<string | null>(null);
  const [confirmReiniciar, setConfirmReiniciar] = useState<boolean>(false);
  const [confirmEliminarId, setConfirmEliminarId] = useState<string | null>(null);
  const [guardandoPedido, setGuardandoPedido] = useState<boolean>(false);

  // 4. Selección múltiple y barra flotante de acciones en lote
  const [selectedIds, setSelectedIds] = useState<Set<string>>(() => new Set<string>());
  const [lastSelectedId, setLastSelectedId] = useState<string | null>(null);
  const [batchPanel, setBatchPanel] = useState<BatchPanelActivo>(null);
  const [batchTransporte, setBatchTransporte] = useState<string>('');
  const [batchVendedor, setBatchVendedor] = useState<string>('Juancho');
  const [pendingUndoDelete, setPendingUndoDelete] = useState<PendingUndoDelete | null>(null);

  // 5. Estado de Vista 2 ("Nuevo pedido")
  const [form, setForm] = useState<FormNuevoPedido>(() => crearFormInicial());
  const [lineasBorrador, setLineasBorrador] = useState<LineaPedido[]>([]);
  const [advertenciasExcel, setAdvertenciasExcel] = useState<AdvertenciaExcel[]>([]);
  const [cantidadesManuales, setCantidadesManuales] = useState<Record<string, number>>({});
  const [nombreArchivoCargado, setNombreArchivoCargado] = useState<string>('');
  const [modoCargaManual, setModoCargaManual] = useState<boolean>(false);
  const [busquedaCatalogo, setBusquedaCatalogo] = useState<string>('');
  const [arrastrandoArchivo, setArrastrandoArchivo] = useState<boolean>(false);
  const [resolucionesManual, setResolucionesManual] = useState<Record<string, { q0: string; q1: string }>>({});

  // 6. Refs de accesibilidad, teclado, archivos y borrado diferido
  const pendingDeleteIdsRef = useRef<Set<string>>(new Set<string>());
  const toastTimerRef = useRef<number | null>(null);
  const searchInputRef = useRef<HTMLInputElement | null>(null);
  const headerCheckboxRef = useRef<HTMLInputElement | null>(null);
  const fileInputListaRef = useRef<HTMLInputElement | null>(null);
  const fileInputGeneralRef = useRef<HTMLInputElement | null>(null);

  // 7. Derivados de sesión (antes de los efectos que los leen)
  const sesionValida = Boolean(sesion && sesion.nombre.trim().length > 0);
  const esAdmin = sesion?.rol === 'Administración';

  // ============================================================================
  // EFECTOS DE HIDRATACIÓN, EXCEL, FIREBASE Y ATAJOS
  // ============================================================================

  useEffect(() => {
    setMounted(true);
    let activo = true;
    void import('xlsx')
      .then((mod) => {
        if (activo) setXLSX(mod as unknown as SheetJSModule);
      })
      .catch(() => {
        // Reintento silencioso si falla carga diferida
      });
    return () => {
      activo = false;
    };
  }, []);

  // Respaldo local híbrido al montar (filtros persistidos + modo local sin nube)
  useEffect(() => {
    if (!mounted) return;
    const savedFilters = cargarFiltrosLocales();
    if (savedFilters.busqueda) setBusqueda(savedFilters.busqueda);
    if (savedFilters.filtro !== 'todos') setFiltro(savedFilters.filtro);

    if (!cloudDb) {
      const cfgRaw = leerJsonLocal(LS_KEYS.config);
      if (typeof cfgRaw === 'object' && cfgRaw !== null) {
        const c = cfgRaw as Record<string, unknown>;
        const adminPinHash = readString(c, 'adminPinHash');
        if (adminPinHash) {
          setConfigEquipo({
            adminPinHash,
            ventasPinHash: readString(c, 'ventasPinHash'),
            updatedAt: readNumber(c, 'updatedAt', 1),
            updatedBy: readString(c, 'updatedBy', 'Administración')
          });
        }
      }
      const sesRaw = leerJsonLocal(LS_KEYS.sesion);
      if (typeof sesRaw === 'object' && sesRaw !== null) {
        const s = sesRaw as Record<string, unknown>;
        const nombre = readString(s, 'nombre');
        if (nombre) {
          const rol: RolEquipo =
            readString(s, 'rol') === 'Administración' ? 'Administración' : 'Ventas / Depósito';
          setSesion({ nombre, rol, configVersion: readNumber(s, 'configVersion', 0) });
        }
      }
      setPedidos(cargarPedidosLocales());
      setCargandoNube(false);
    }
  }, [mounted]);

  // Autenticación Firebase y suscripción en tiempo real (solo si hay configuración cloud)
  useEffect(() => {
    if (!mounted || !cloudDb || !firebaseApp) return;
    const auth = getAuth(firebaseApp);
    const unsubAuth = onAuthStateChanged(auth, (fbUser) => {
      if (fbUser) {
        setUser(fbUser);
      } else {
        void signInAnonymously(auth).catch(() => {
          setErrorNube(true);
          setCargandoNube(false);
        });
      }
    });
    return () => unsubAuth();
  }, [mounted]);

  useEffect(() => {
    if (!mounted || !cloudDb || !user) return;
    const db = cloudDb;

    const unsubConfig = onSnapshot(
      getTeamConfigDoc(db),
      (snap) => {
        if (snap.exists()) {
          const c = snap.data() as Record<string, unknown>;
          const adminPinHash = readString(c, 'adminPinHash');
          if (adminPinHash) {
            setConfigEquipo({
              adminPinHash,
              ventasPinHash: readString(c, 'ventasPinHash'),
              updatedAt: readNumber(c, 'updatedAt', 1),
              updatedBy: readString(c, 'updatedBy', 'Administración')
            });
          }
        }
      },
      () => {
        setErrorNube(true);
      }
    );

    const unsubSession = onSnapshot(
      getUserSessionDoc(db, user.uid),
      (snap) => {
        if (snap.exists()) {
          const s = snap.data() as Record<string, unknown>;
          const nombre = readString(s, 'nombre');
          if (nombre) {
            const rol: RolEquipo =
              readString(s, 'rol') === 'Administración' ? 'Administración' : 'Ventas / Depósito';
            setSesion({ nombre, rol, configVersion: readNumber(s, 'configVersion', 0) });
          } else {
            setSesion(null);
          }
        }
      },
      () => {
        // Mantener sesión local si falla lectura de sesión
      }
    );

    const unsubOrders = onSnapshot(
      getPublicOrdersCollection(db),
      (snapshot) => {
        const cargadosMap = new Map<string, Pedido>();
        let invalidos = 0;
        snapshot.forEach((docSnap) => {
          if (pendingDeleteIdsRef.current.has(docSnap.id)) return;
          const data = docSnap.data() as Record<string, unknown>;
          const pedido = deserializePedidoFromCloud(docSnap.id, data);
          if (pedido) {
            cargadosMap.set(pedido.id, pedido);
          } else {
            invalidos += 1;
          }
        });

        const cargados = Array.from(cargadosMap.values());
        cargados.sort((a, b) => {
          const diffTime = (b.createdAt ?? 0) - (a.createdAt ?? 0);
          if (diffTime !== 0) return diffTime;
          return (parseInt(b.numeroPedido, 10) || 0) - (parseInt(a.numeroPedido, 10) || 0);
        });

        setPedidos(cargados);
        guardarJsonLocal(LS_KEYS.pedidos, cargados);
        setPedidosInvalidosNube(invalidos);
        setCargandoNube(false);
        setErrorNube(false);
      },
      () => {
        setErrorNube(true);
        setCargandoNube(false);
      }
    );

    return () => {
      unsubConfig();
      unsubSession();
      unsubOrders();
    };
  }, [mounted, user]);

  // Persistencia automática de filtros y respaldo local de pedidos
  useEffect(() => {
    if (!mounted) return;
    guardarJsonLocal(LS_KEYS.filtros, { busqueda, filtro });
  }, [mounted, busqueda, filtro]);

  useEffect(() => {
    if (!mounted || cloudDb) return;
    guardarJsonLocal(LS_KEYS.pedidos, pedidos);
  }, [mounted, pedidos]);

  // Limpia IDs seleccionados que ya no existen
  useEffect(() => {
    setSelectedIds((prev) => {
      if (prev.size === 0) return prev;
      const existentes = new Set(pedidos.map((p) => p.id));
      let cambio = false;
      const next = new Set<string>();
      prev.forEach((id) => {
        if (existentes.has(id)) next.add(id);
        else cambio = true;
      });
      return cambio ? next : prev;
    });
  }, [pedidos]);

  useEffect(() => {
    if (selectedIds.size === 0) {
      setBatchPanel(null);
    }
  }, [selectedIds.size]);

  // Atajos de teclado globales (/, Cmd/Ctrl+K, N, Escape)
  useEffect(() => {
    if (!sesionValida) return;

    const handleKeyDown = (e: KeyboardEvent): void => {
      const target = e.target as HTMLElement | null;
      const isEditing = Boolean(
        target &&
          (target.tagName === 'INPUT' ||
            target.tagName === 'TEXTAREA' ||
            target.tagName === 'SELECT' ||
            target.isContentEditable)
      );

      if (e.key === 'Escape') {
        if (modalEquipoAbierto) {
          setModalEquipoAbierto(false);
          return;
        }
        if (drawerPedidoId) {
          setDrawerPedidoId(null);
          setMostrarHojaImprimible(false);
          return;
        }
        if (batchPanel) {
          setBatchPanel(null);
          return;
        }
        if (selectedIds.size > 0) {
          setSelectedIds(new Set<string>());
          setLastSelectedId(null);
          return;
        }
        return;
      }

      if ((e.key.toLowerCase() === 'k' && (e.metaKey || e.ctrlKey)) || (e.key === '/' && !isEditing)) {
        e.preventDefault();
        if (vista !== 'pedidos') setVista('pedidos');
        window.setTimeout(() => {
          searchInputRef.current?.focus();
          searchInputRef.current?.select();
        }, 20);
        return;
      }

      if (
        !isEditing &&
        !e.metaKey &&
        !e.ctrlKey &&
        !e.altKey &&
        e.key.toLowerCase() === 'n' &&
        !modalEquipoAbierto &&
        !drawerPedidoId
      ) {
        e.preventDefault();
        setVista('nuevo');
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [sesionValida, vista, modalEquipoAbierto, drawerPedidoId, batchPanel, selectedIds.size]);

  // ============================================================================
  // CÁLCULOS MEMOIZADOS Y DERIVADOS
  // ============================================================================

  const resumenFinanciero = useMemo<ResumenFinanciero>(
    () => calcularResumenFinanciero(pedidos),
    [pedidos]
  );

  const conteosFiltro = useMemo(
    () => ({
      todos: pedidos.length,
      falta_talle: pedidos.filter((p) => p.faltantesVariante > 0).length,
      cheques: pedidos.filter((p) => p.medioPago === 'Cheque' && !p.chequeAcreditado).length,
      falta_pago: pedidos.filter((p) => p.estadoPago === 'Falta pago').length,
      listos: pedidos.filter((p) => p.estadoArmado === 'Listo').length
    }),
    [pedidos]
  );

  const pedidosFiltrados = useMemo<Pedido[]>(() => {
    const q = busqueda.trim().toLowerCase();
    return pedidos.filter((p) => {
      if (filtro === 'falta_talle' && p.faltantesVariante === 0) return false;
      if (filtro === 'cheques' && !(p.medioPago === 'Cheque' && !p.chequeAcreditado)) return false;
      if (filtro === 'falta_pago' && p.estadoPago !== 'Falta pago') return false;
      if (filtro === 'listos' && p.estadoArmado !== 'Listo') return false;

      if (!q) return true;
      return (
        p.cliente.toLowerCase().includes(q) ||
        p.vendedor.toLowerCase().includes(q) ||
        p.numeroPedido.toLowerCase().includes(q) ||
        p.detalleResumen.toLowerCase().includes(q)
      );
    });
  }, [pedidos, busqueda, filtro]);

  const pedidosSeleccionados = useMemo<Pedido[]>(
    () => pedidos.filter((p) => selectedIds.has(p.id)),
    [pedidos, selectedIds]
  );

  const totalSeleccionado = useMemo<number>(
    () => round2(pedidosSeleccionados.reduce((acc, p) => acc + p.totalFacturado, 0)),
    [pedidosSeleccionados]
  );

  const visiblesSeleccionadosCount = useMemo<number>(
    () => pedidosFiltrados.reduce((acc, p) => acc + (selectedIds.has(p.id) ? 1 : 0), 0),
    [pedidosFiltrados, selectedIds]
  );

  const todosVisiblesSeleccionados =
    pedidosFiltrados.length > 0 && visiblesSeleccionadosCount === pedidosFiltrados.length;
  const seleccionParcialVisible =
    visiblesSeleccionadosCount > 0 && visiblesSeleccionadosCount < pedidosFiltrados.length;

  useEffect(() => {
    if (headerCheckboxRef.current) {
      headerCheckboxRef.current.indeterminate = seleccionParcialVisible;
    }
  }, [seleccionParcialVisible]);

  const pedidoDrawer = useMemo<Pedido | null>(
    () => (drawerPedidoId ? pedidos.find((p) => p.id === drawerPedidoId) ?? null : null),
    [pedidos, drawerPedidoId]
  );

  // Sincroniza el borrador local del Drawer solo cuando cambia el pedido abierto
  useEffect(() => {
    setConfirmEliminarId(null);
    if (!pedidoDrawer) {
      setDrawerDraft(null);
      return;
    }
    setDrawerDraft({
      cliente: pedidoDrawer.cliente,
      vendedor: pedidoDrawer.vendedor,
      whatsapp: pedidoDrawer.whatsapp,
      despachoPrevisto: pedidoDrawer.despachoPrevisto,
      descuentoPorc: String(pedidoDrawer.descuentoPorc),
      montoAbonado: String(pedidoDrawer.montoAbonado),
      montoCheque: String(pedidoDrawer.montoCheque),
      fechaRecepcionCheque: pedidoDrawer.fechaRecepcionCheque,
      fechaAcreditacionCheque: pedidoDrawer.fechaAcreditacionCheque,
      detalleCheque: pedidoDrawer.detalleCheque,
      transporte: pedidoDrawer.transporte,
      nroSeguimiento: pedidoDrawer.nroSeguimiento,
      observaciones: pedidoDrawer.observaciones
    });
  }, [drawerPedidoId]);

  const totalesBorrador = useMemo(() => {
    const subtotalNeto = round2(
      lineasBorrador.reduce((acc, l) => acc + round2(l.precioUnitarioNeto * l.cantidadTotal), 0)
    );
    const desc = Math.max(0, Math.min(100, Number.isFinite(form.descuentoPorc) ? form.descuentoPorc : 0));
    const montoDescuento = round2((subtotalNeto * desc) / 100);
    const baseImponible = round2(subtotalNeto - montoDescuento);
    const iva21 = round2(baseImponible * 0.21);
    const totalFacturado = round2(baseImponible + iva21);
    const faltantes = lineasBorrador.filter((l) => l.requiereRevision || l.estado !== 'OK').length;
    const totalUnidades = lineasBorrador.reduce((acc, l) => acc + l.cantidadTotal, 0);
    return {
      subtotalNeto,
      montoDescuento,
      baseImponible,
      iva21,
      totalFacturado,
      faltantes,
      totalUnidades
    };
  }, [lineasBorrador, form.descuentoPorc]);

  const hayBorradorSinGuardar =
    lineasBorrador.length > 0 || form.cliente.trim().length > 0 || form.observaciones.trim().length > 0;

  const proximoNumeroPedido = useMemo<string>(() => {
    const maxNum = pedidos.reduce((acc, p) => {
      const n = parseInt(p.numeroPedido.replace(/\D/g, ''), 10);
      return Number.isFinite(n) && n > acc ? n : acc;
    }, 0);
    return String(maxNum + 1).padStart(4, '0');
  }, [pedidos]);

  // ============================================================================
  // HELPERS Y HANDLERS (TODOS DECLARADOS ANTES DE EJECUTAR LOTE / JSX)
  // ============================================================================

  const notify = (texto: string, tone: ToneStatus = 'neutral'): void => {
    if (toastTimerRef.current !== null) {
      window.clearTimeout(toastTimerRef.current);
    }
    setToast({ texto, tone });
    toastTimerRef.current = window.setTimeout(() => {
      setToast(null);
      toastTimerRef.current = null;
    }, 3400);
  };

  const crearEventoHistorial = (accion: string): HistorialEvento => ({
    fechaHora: formatNowDateTime(),
    usuario: sesion?.nombre ?? 'Equipo',
    rol: sesion?.rol ?? 'Ventas / Depósito',
    accion
  });

  // Declarado ANTES de ejecutarEdicionLote para evitar cualquier TS2448
  const actualizarPedido = (pedidoId: string, patch: Partial<PedidoBase>, motivo: string): void => {
    const evento = crearEventoHistorial(motivo);
    const ahora = Date.now();
    const pedidoActual = pedidos.find((p) => p.id === pedidoId);
    const historialNuevo: HistorialEvento[] = pedidoActual ? [evento, ...pedidoActual.historial] : [evento];

    setPedidos((prev) =>
      prev.map((p) => {
        if (p.id !== pedidoId) return p;
        return recalcularPedido({
          ...p,
          ...patch,
          updatedAt: ahora,
          historial: [evento, ...p.historial]
        });
      })
    );

    if (cloudDb && user) {
      void patchOrderInCloud(cloudDb, pedidoId, patch, historialNuevo, ahora).catch(() => {
        notify('Cambio guardado localmente; reintentando sincronización.', 'amber');
      });
    }
  };

  const confirmarAcreditacion = (pedido: Pedido): void => {
    if (!esAdmin) {
      notify('Solo Administración puede confirmar acreditación de cheques.', 'amber');
      return;
    }
    const hoy = getTodayISO();
    const nota = `Cheque acreditado el ${formatShortDate(hoy)}`;
    actualizarPedido(
      pedido.id,
      {
        medioPago: 'Transferencia / Efectivo',
        chequeAcreditado: true,
        montoAbonado: pedido.totalFacturado,
        montoCheque: 0,
        observaciones: [pedido.observaciones, nota].filter(Boolean).join(' · ')
      },
      `Confirmó plata acreditada (${formatARS(pedido.totalFacturado)}) → Pagado`
    );
    notify(`Pedido #${pedido.numeroPedido} pasado a Pagado.`, 'green');
  };

  const eliminarPedidoIndividual = (pedido: Pedido): void => {
    if (!esAdmin) {
      notify('Solo Administración puede eliminar pedidos.', 'amber');
      return;
    }
    const id = pedido.id;
    pendingDeleteIdsRef.current.add(id);
    setPedidos((prev) => prev.filter((p) => p.id !== id));
    if (drawerPedidoId === id) {
      setDrawerPedidoId(null);
      setMostrarHojaImprimible(false);
    }
    setConfirmEliminarId(null);

    if (cloudDb && user) {
      void clearAllOrdersInCloud(cloudDb, [id]).finally(() => {
        pendingDeleteIdsRef.current.delete(id);
      });
    } else {
      pendingDeleteIdsRef.current.delete(id);
    }
    notify(`Pedido #${pedido.numeroPedido} eliminado.`, 'neutral');
  };

  const handleToggleSelectAllVisible = (): void => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (todosVisiblesSeleccionados) {
        for (const p of pedidosFiltrados) next.delete(p.id);
      } else {
        for (const p of pedidosFiltrados) next.add(p.id);
      }
      return next;
    });
  };

  const handleToggleSelectRow = (id: string, shiftKey: boolean): void => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (shiftKey && lastSelectedId) {
        const idxA = pedidosFiltrados.findIndex((p) => p.id === lastSelectedId);
        const idxB = pedidosFiltrados.findIndex((p) => p.id === id);
        if (idxA !== -1 && idxB !== -1) {
          const start = Math.min(idxA, idxB);
          const end = Math.max(idxA, idxB);
          const marcar = !prev.has(id);
          for (let i = start; i <= end; i++) {
            const rowId = pedidosFiltrados[i].id;
            if (marcar) next.add(rowId);
            else next.delete(rowId);
          }
          return next;
        }
      }
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
    setLastSelectedId(id);
  };

  const limpiarSeleccion = (): void => {
    setSelectedIds(new Set<string>());
    setLastSelectedId(null);
    setBatchPanel(null);
  };

  // Motor de edición masiva segura
  const ejecutarEdicionLote = (
    updater: (p: Pedido) => { patch: Partial<PedidoBase>; motivo: string } | null
  ): { actualizados: number; omitidos: Pedido[] } => {
    const omitidos: Pedido[] = [];
    let actualizados = 0;

    for (const p of pedidosSeleccionados) {
      const res = updater(p);
      if (!res) {
        omitidos.push(p);
        continue;
      }
      actualizarPedido(p.id, res.patch, res.motivo);
      actualizados += 1;
    }

    return { actualizados, omitidos };
  };

  const handleBatchArmado = (destino: 'En armado' | 'Listo'): void => {
    const { actualizados, omitidos } = ejecutarEdicionLote((p) => {
      if (destino === 'Listo' && p.faltantesVariante > 0) return null;
      return {
        patch: { estadoArmado: destino },
        motivo: `Acción en lote: pasó armado a ${destino}`
      };
    });

    setBatchPanel(null);
    if (omitidos.length > 0 && actualizados > 0) {
      notify(
        `${actualizados} pasados a Listo · ${omitidos.length} bloqueados por falta de talle (${omitidos
          .map((o) => `#${o.numeroPedido}`)
          .join(', ')})`,
        'amber'
      );
    } else if (omitidos.length > 0 && actualizados === 0) {
      notify(
        `Ninguno pudo pasar a Listo: falta definir talle o color en ${omitidos
          .map((o) => `#${o.numeroPedido}`)
          .join(', ')}`,
        'red'
      );
    } else {
      notify(`${actualizados} ${actualizados === 1 ? 'pedido pasado' : 'pedidos pasados'} a ${destino}`, 'green');
    }
  };

  const handleBatchEnvio = (destino: 'Despachado' | 'Entregado'): void => {
    const transp = batchTransporte.trim();
    const { actualizados } = ejecutarEdicionLote(() => ({
      patch: {
        estadoEnvio: destino,
        ...(transp ? { transporte: transp } : {})
      },
      motivo: transp
        ? `Acción en lote: ${destino} por ${transp}`
        : `Acción en lote: envío marcado como ${destino}`
    }));

    setBatchPanel(null);
    setBatchTransporte('');
    notify(
      transp
        ? `${actualizados} pedidos marcados como ${destino} (${transp})`
        : `${actualizados} pedidos marcados como ${destino}`,
      'green'
    );
  };

  const handleBatchPago = (accion: 'Pagado' | 'Falta pago'): void => {
    if (!esAdmin) {
      notify('Solo Administración puede modificar pagos en lote.', 'amber');
      return;
    }
    const hoy = getTodayISO();
    const { actualizados } = ejecutarEdicionLote((p) => {
      if (accion === 'Pagado') {
        const nota =
          p.medioPago === 'Cheque'
            ? `Cheque acreditado en lote el ${formatShortDate(hoy)}`
            : `Cobro confirmado en lote el ${formatShortDate(hoy)}`;
        return {
          patch: {
            medioPago: 'Transferencia / Efectivo',
            chequeAcreditado: true,
            montoAbonado: p.totalFacturado,
            montoCheque: 0,
            observaciones: [p.observaciones, nota].filter(Boolean).join(' · ')
          },
          motivo: `Acción en lote: marcado como Pagado (${formatARS(p.totalFacturado)})`
        };
      }
      return {
        patch: {
          medioPago: 'Sin pagar',
          chequeAcreditado: false,
          montoAbonado: 0,
          montoCheque: 0
        },
        motivo: 'Acción en lote: pasado a Falta pago'
      };
    });

    setBatchPanel(null);
    notify(
      `${actualizados} ${actualizados === 1 ? 'pedido actualizado' : 'pedidos actualizados'} a ${accion}`,
      accion === 'Pagado' ? 'green' : 'neutral'
    );
  };

  const handleBatchVendedor = (): void => {
    if (!esAdmin) {
      notify('Solo Administración puede reasignar vendedores en lote.', 'amber');
      return;
    }
    const vend = batchVendedor.trim();
    if (!vend) return;
    const { actualizados } = ejecutarEdicionLote(() => ({
      patch: { vendedor: vend },
      motivo: `Acción en lote: vendedor asignado a ${vend}`
    }));
    setBatchPanel(null);
    notify(`${actualizados} pedidos asignados a ${vend}`, 'green');
  };

  const purgarEliminacionPendiente = (pendiente: PendingUndoDelete): void => {
    window.clearTimeout(pendiente.timerId);
    const ids = pendiente.pedidos.map((p) => p.id);
    if (cloudDb && user && ids.length > 0) {
      void clearAllOrdersInCloud(cloudDb, ids).finally(() => {
        for (const id of ids) pendingDeleteIdsRef.current.delete(id);
      });
    } else {
      for (const id of ids) pendingDeleteIdsRef.current.delete(id);
    }
  };

  const handleConfirmarEliminarLote = (): void => {
    if (!esAdmin || pedidosSeleccionados.length === 0) return;

    if (pendingUndoDelete) {
      purgarEliminacionPendiente(pendingUndoDelete);
    }

    const aEliminar = [...pedidosSeleccionados];
    const idsSet = new Set(aEliminar.map((p) => p.id));
    idsSet.forEach((id) => pendingDeleteIdsRef.current.add(id));

    setPedidos((prev) => prev.filter((p) => !idsSet.has(p.id)));
    if (drawerPedidoId && idsSet.has(drawerPedidoId)) {
      setDrawerPedidoId(null);
      setMostrarHojaImprimible(false);
    }
    limpiarSeleccion();

    const timerId = window.setTimeout(() => {
      const ids = aEliminar.map((p) => p.id);
      if (cloudDb && user && ids.length > 0) {
        void clearAllOrdersInCloud(cloudDb, ids)
          .catch(() => {
            notify('No se pudo completar el borrado en la nube.', 'red');
          })
          .finally(() => {
            for (const id of ids) pendingDeleteIdsRef.current.delete(id);
          });
      } else {
        for (const id of ids) pendingDeleteIdsRef.current.delete(id);
      }
      setPendingUndoDelete((curr) => (curr && curr.timerId === timerId ? null : curr));
    }, 6000);

    setPendingUndoDelete({ pedidos: aEliminar, timerId });
  };

  const handleDeshacerEliminacion = (): void => {
    if (!pendingUndoDelete) return;
    window.clearTimeout(pendingUndoDelete.timerId);
    const restaurar = pendingUndoDelete.pedidos;
    for (const p of restaurar) pendingDeleteIdsRef.current.delete(p.id);

    setPedidos((prev) => {
      const map = new Map<string, Pedido>();
      for (const p of [...restaurar, ...prev]) map.set(p.id, p);
      return Array.from(map.values()).sort((a, b) => {
        const diffTime = (b.createdAt ?? 0) - (a.createdAt ?? 0);
        if (diffTime !== 0) return diffTime;
        return (parseInt(b.numeroPedido, 10) || 0) - (parseInt(a.numeroPedido, 10) || 0);
      });
    });
    setPendingUndoDelete(null);
    notify(
      `${restaurar.length} ${restaurar.length === 1 ? 'pedido restaurado' : 'pedidos restaurados'}.`,
      'green'
    );
  };

  const handleDuplicarComoNuevoPedido = (pedido: Pedido): void => {
    const hoy = getTodayISO();
    setForm({
      ...crearFormInicial(),
      cliente: pedido.cliente,
      vendedor: pedido.vendedor,
      whatsapp: pedido.whatsapp,
      fechaPedido: hoy,
      despachoPrevisto: addDaysISO(hoy, 5),
      pagoTipo: 'Sin pagar'
    });
    setLineasBorrador([]);
    setAdvertenciasExcel([]);
    setCantidadesManuales({});
    setNombreArchivoCargado('');
    setModoCargaManual(false);
    setDrawerPedidoId(null);
    setMostrarHojaImprimible(false);
    setVista('nuevo');
    notify(`Datos de ${pedido.cliente} copiados para un nuevo pedido.`, 'green');
  };

  // Resolución manual inline (tanto para Vista 2 borrador como para Drawer)
  const aplicarResolucionLinea = (linea: LineaPedido, q0: number, q1: number): LineaPedido | null => {
    const prod = CATALOGO_BY_GRUPO.get(linea.idGrupo);
    if (!prod || prod.variantes.length < 2) return null;
    if (q0 < 0 || q1 < 0 || !Number.isInteger(q0) || !Number.isInteger(q1)) return null;
    if (q0 + q1 !== linea.cantidadTotal) return null;

    const [v0, v1] = prod.variantes;
    return {
      ...linea,
      estado: 'OK',
      requiereRevision: false,
      motivoRevision: undefined,
      desglose: [
        { codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: q0 },
        { codigo: v1.codigo, etiqueta: v1.etiqueta, cantidad: q1 }
      ]
    };
  };

  const handleResolverEnBorrador = (idLinea: string, q0: number, q1: number): void => {
    const lineaObj = lineasBorrador.find((l) => l.idLinea === idLinea);
    if (!lineaObj) return;
    const resuelta = aplicarResolucionLinea(lineaObj, q0, q1);
    if (!resuelta) {
      notify(`La suma debe ser exactamente ${lineaObj.cantidadTotal} u.`, 'red');
      return;
    }
    setLineasBorrador((prev) => prev.map((l) => (l.idLinea === idLinea ? resuelta : l)));
    notify(`${lineaObj.nombreProducto} confirmado.`, 'green');
  };

  const handleResolverEnDrawer = (pedido: Pedido, idLinea: string, q0: number, q1: number): void => {
    const lineaObj = pedido.lineas.find((l) => l.idLinea === idLinea);
    if (!lineaObj) return;
    const resuelta = aplicarResolucionLinea(lineaObj, q0, q1);
    if (!resuelta) {
      notify(`La suma debe ser exactamente ${lineaObj.cantidadTotal} u.`, 'red');
      return;
    }
    const nuevasLineas = pedido.lineas.map((l) => (l.idLinea === idLinea ? resuelta : l));
    actualizarPedido(
      pedido.id,
      { lineas: nuevasLineas },
      `Definió talle/color en ${lineaObj.nombreProducto} (${q0} / ${q1})`
    );
    notify(`${lineaObj.nombreProducto} confirmado para depósito.`, 'green');
  };

  // Sincronización fuente única de verdad entre carga manual y líneas del borrador
  const sincronizarManualConBorrador = (nextMap: Record<string, number>): void => {
    setLineasBorrador((prev) => {
      const prevByGrupo = new Map<string, LineaPedido>();
      for (const l of prev) prevByGrupo.set(l.idGrupo, l);

      const resultado: LineaPedido[] = [];
      for (const prod of CATALOGO_MIBI) {
        const qty = nextMap[prod.idGrupo] ?? 0;
        if (qty <= 0) continue;
        const existente = prevByGrupo.get(prod.idGrupo);
        if (existente && existente.cantidadTotal === qty) {
          resultado.push(existente);
        } else {
          resultado.push(resolverLinea(prod, qty, existente?.notaCliente ?? ''));
        }
      }
      return resultado;
    });
  };

  const handleCambioCantidadManual = (prod: ProductoCatalogo, valorStr: string): void => {
    const parsed = valorStr === '' ? 0 : Math.max(0, Math.floor(Number(valorStr) || 0));
    setCantidadesManuales((prev) => {
      const next = { ...prev };
      if (parsed <= 0) delete next[prod.idGrupo];
      else next[prod.idGrupo] = parsed;
      sincronizarManualConBorrador(next);
      return next;
    });
  };

  const handleEliminarLineaBorrador = (idGrupo: string): void => {
    setLineasBorrador((prev) => prev.filter((l) => l.idGrupo !== idGrupo));
    setCantidadesManuales((prev) => {
      const next = { ...prev };
      delete next[idGrupo];
      return next;
    });
  };

  const handleArchivoLista006 = async (file: File): Promise<void> => {
    if (!XLSX) {
      notify('Preparando lector Excel, intentá en un segundo.', 'amber');
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const res = parsearExcelLista006(buffer, XLSX, file.name);
      if (!res.ok) {
        notify(res.errorFatal ?? 'Archivo inválido.', 'red');
        return;
      }
      const mapaManual: Record<string, number> = {};
      for (const l of res.lineas) {
        mapaManual[l.idGrupo] = l.cantidadTotal;
      }
      setLineasBorrador(res.lineas);
      setCantidadesManuales(mapaManual);
      setAdvertenciasExcel(res.advertencias);
      setNombreArchivoCargado(file.name);
      if (res.advertencias.length > 0) {
        notify(`Se detectaron ${res.advertencias.length} fila(s) con error para revisar antes de guardar.`, 'red');
      } else {
        notify(`${res.lineas.length} artículos leídos de ${file.name}.`, 'green');
      }
    } catch {
      notify('No se pudo leer el archivo seleccionado.', 'red');
    }
  };

  const handleCargarEjemploLista006 = (): void => {
    const ejemplo = crearLineasEjemploLista006();
    const mapa: Record<string, number> = {};
    for (const l of ejemplo) mapa[l.idGrupo] = l.cantidadTotal;
    setForm((prev) => ({
      ...prev,
      cliente: prev.cliente || 'Casa Lorenzo (Ejemplo Lista 006)',
      descuentoPorc: 0
    }));
    setLineasBorrador(ejemplo);
    setCantidadesManuales(mapa);
    setAdvertenciasExcel([]);
    setNombreArchivoCargado('Ejemplo real LISTA N° 006');
    notify('Planilla de ejemplo Lista N° 006 cargada ($3.560.673,05 total).', 'green');
  };

  const handleGuardarNuevoPedido = async (): Promise<void> => {
    if (!form.cliente.trim()) {
      notify('Ingresá el nombre del cliente o negocio.', 'amber');
      return;
    }
    if (lineasBorrador.length === 0) {
      notify('Cargá al menos un artículo desde Excel o en carga manual.', 'amber');
      return;
    }
    if (advertenciasExcel.length > 0) {
      notify('Confirmá o descartá las alertas de Excel antes de guardar.', 'red');
      return;
    }

    setGuardandoPedido(true);
    try {
      const totalEst = totalesBorrador.totalFacturado;
      const medioPagoEfectivo: MedioPago = !esAdmin
        ? 'Sin pagar'
        : form.pagoTipo;

      let montoAbonado = 0;
      let montoCheque = 0;

      if (esAdmin && medioPagoEfectivo === 'Transferencia / Efectivo') {
        const val = parseFloat(form.montoAbonadoInput.replace(',', '.'));
        montoAbonado = Number.isFinite(val) && val > 0 ? Math.min(totalEst, round2(val)) : totalEst;
      } else if (esAdmin && medioPagoEfectivo === 'Cheque') {
        const valAbon = parseFloat(form.montoAbonadoInput.replace(',', '.'));
        montoAbonado = Number.isFinite(valAbon) && valAbon > 0 ? Math.min(totalEst, round2(valAbon)) : 0;
        const valCheq = parseFloat(form.montoChequeInput.replace(',', '.'));
        montoCheque =
          Number.isFinite(valCheq) && valCheq > 0
            ? Math.min(round2(totalEst - montoAbonado), round2(valCheq))
            : round2(totalEst - montoAbonado);
      }

      const id = `ped_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
      const ahora = Date.now();
      const nuevo = recalcularPedido({
        id,
        numeroPedido: proximoNumeroPedido,
        cliente: form.cliente.trim(),
        vendedor: form.vendedor,
        whatsapp: form.whatsapp.trim(),
        fechaPedido: form.fechaPedido || getTodayISO(),
        despachoPrevisto: form.despachoPrevisto || addDaysISO(getTodayISO(), 5),
        lineas: lineasBorrador,
        descuentoPorc: form.descuentoPorc,
        medioPago: medioPagoEfectivo,
        montoAbonado,
        montoCheque,
        chequeAcreditado: false,
        fechaRecepcionCheque: medioPagoEfectivo === 'Cheque' ? form.fechaRecepcionCheque : '',
        fechaAcreditacionCheque: medioPagoEfectivo === 'Cheque' ? form.fechaAcreditacionCheque : '',
        detalleCheque: medioPagoEfectivo === 'Cheque' ? form.detalleCheque.trim() : '',
        estadoArmado: 'Por armar',
        estadoEnvio: 'Pendiente',
        transporte: '',
        nroSeguimiento: '',
        observaciones: form.observaciones.trim(),
        createdAt: ahora,
        updatedAt: ahora,
        historial: [crearEventoHistorial(`Creó el pedido #${proximoNumeroPedido}`)]
      });

      setPedidos((prev) => [nuevo, ...prev]);
      if (cloudDb && user) {
        await saveOrderInCloud(cloudDb, nuevo);
      }

      setForm(crearFormInicial());
      setLineasBorrador([]);
      setAdvertenciasExcel([]);
      setCantidadesManuales({});
      setNombreArchivoCargado('');
      setModoCargaManual(false);
      setResolucionesManual({});
      setVista('pedidos');
      notify(`Pedido #${nuevo.numeroPedido} guardado (${formatARS(nuevo.totalFacturado)}).`, 'green');
    } catch {
      notify('No se pudo guardar el pedido.', 'red');
    } finally {
      setGuardandoPedido(false);
    }
  };

  const handleImportarPlanillaGeneral = async (file: File): Promise<void> => {
    if (!XLSX) {
      notify('Preparando lector Excel, intentá en un segundo.', 'amber');
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const res = importarPlanillaGeneral(buffer, XLSX, file.name, pedidos);
      if (!res.ok) {
        notify(res.error ?? 'Error al importar planilla.', 'red');
        return;
      }
      if (res.nuevos.length === 0) {
        notify(
          res.duplicadosOmitidos > 0
            ? `Sin pedidos nuevos (${res.duplicadosOmitidos} duplicados omitidos).`
            : 'No se encontraron pedidos válidos en la planilla.',
          'amber'
        );
        return;
      }
      setPedidos((prev) => [...res.nuevos, ...prev]);
      if (cloudDb && user) {
        for (const p of res.nuevos) {
          void saveOrderInCloud(cloudDb, p);
        }
      }
      notify(
        `Importados ${res.nuevos.length} pedidos (${res.duplicadosOmitidos} duplicados omitidos).`,
        'green'
      );
    } catch {
      notify('Error al procesar la planilla general.', 'red');
    }
  };

  // Autenticación de equipo y gestión de claves
  const handleConfigurarAccesoInicial = async (e: React.FormEvent<HTMLFormElement>): Promise<void> => {
    e.preventDefault();
    setErrorIngreso('');
    const nombre = nombreIngreso.trim();
    const pAdmin = pinAdminNuevo.trim();
    const pVentas = pinVentasNuevo.trim();

    if (!nombre) {
      setErrorIngreso('Ingresá tu nombre.');
      return;
    }
    if (pAdmin.length < 4) {
      setErrorIngreso('La clave de Administración debe tener al menos 4 caracteres.');
      return;
    }
    if (pVentas && pVentas === pAdmin) {
      setErrorIngreso('La clave de Ventas / Depósito debe ser distinta a la de Administración.');
      return;
    }

    setGuardandoAcceso(true);
    try {
      const adminPinHash = await hashPin(pAdmin);
      const ventasPinHash = pVentas ? await hashPin(pVentas) : '';
      const nuevaConfig: ConfigEquipo = {
        adminPinHash,
        ventasPinHash,
        updatedAt: Date.now(),
        updatedBy: nombre
      };
      const nuevaSesion: SesionUsuario = {
        nombre,
        rol: 'Administración',
        configVersion: nuevaConfig.updatedAt
      };

      if (cloudDb && user) {
        await saveTeamConfigToCloud(cloudDb, nuevaConfig);
        await saveUserSessionToCloud(cloudDb, user.uid, nuevaSesion);
      } else {
        guardarJsonLocal(LS_KEYS.config, nuevaConfig);
        guardarJsonLocal(LS_KEYS.sesion, nuevaSesion);
      }
      setConfigEquipo(nuevaConfig);
      setSesion(nuevaSesion);
      setPinAdminNuevo('');
      setPinVentasNuevo('');
      notify(`Bienvenido/a, ${nombre}. Acceso activado.`, 'green');
    } catch {
      setErrorIngreso('No se pudo guardar la configuración inicial.');
    } finally {
      setGuardandoAcceso(false);
    }
  };

  const handleIngresarEquipo = async (e: React.FormEvent<HTMLFormElement>): Promise<void> => {
    e.preventDefault();
    setErrorIngreso('');
    if (!configEquipo) return;

    const nombre = nombreIngreso.trim();
    const pin = pinIngreso.trim();
    if (!nombre || !pin) {
      setErrorIngreso('Completá tu nombre y clave.');
      return;
    }

    setGuardandoAcceso(true);
    try {
      const hashIngresado = await hashPin(pin);
      let rolDetectado: RolEquipo | null = null;
      if (hashIngresado === configEquipo.adminPinHash) {
        rolDetectado = 'Administración';
      } else if (configEquipo.ventasPinHash && hashIngresado === configEquipo.ventasPinHash) {
        rolDetectado = 'Ventas / Depósito';
      }

      if (!rolDetectado) {
        setErrorIngreso('Clave incorrecta.');
        return;
      }

      const nuevaSesion: SesionUsuario = {
        nombre,
        rol: rolDetectado,
        configVersion: configEquipo.updatedAt
      };

      if (cloudDb && user) {
        await saveUserSessionToCloud(cloudDb, user.uid, nuevaSesion);
      } else {
        guardarJsonLocal(LS_KEYS.sesion, nuevaSesion);
      }
      setSesion(nuevaSesion);
      setPinIngreso('');
      notify(`Ingresaste como ${nombre} (${rolDetectado}).`, 'green');
    } catch {
      setErrorIngreso('Error al verificar clave.');
    } finally {
      setGuardandoAcceso(false);
    }
  };

  const handleCerrarSesion = async (): Promise<void> => {
    setSesion(null);
    setDrawerPedidoId(null);
    setModalEquipoAbierto(false);
    if (cloudDb && user) {
      try {
        await clearUserSessionInCloud(cloudDb, user.uid);
      } catch {
        // Sesión limpiada localmente aunque falle la red
      }
    } else {
      guardarJsonLocal(LS_KEYS.sesion, null);
    }
    notify('Sesión bloqueada.');
  };

  const handleActualizarClavesEquipo = async (e: React.FormEvent<HTMLFormElement>): Promise<void> => {
    e.preventDefault();
    if (!esAdmin || !configEquipo || !sesion) return;

    const pAdmin = pinAdminModal.trim();
    const pVentas = pinVentasModal.trim();
    if (pAdmin && pAdmin.length < 4) {
      notify('La clave de Administración debe tener mínimo 4 caracteres.', 'amber');
      return;
    }

    setGuardandoAcceso(true);
    try {
      const nuevaConfig: ConfigEquipo = {
        adminPinHash: pAdmin ? await hashPin(pAdmin) : configEquipo.adminPinHash,
        ventasPinHash: pVentas ? await hashPin(pVentas) : configEquipo.ventasPinHash,
        updatedAt: Date.now(),
        updatedBy: sesion.nombre
      };
      const sesionActualizada: SesionUsuario = {
        ...sesion,
        configVersion: nuevaConfig.updatedAt
      };

      if (cloudDb && user) {
        await saveTeamConfigToCloud(cloudDb, nuevaConfig);
        await saveUserSessionToCloud(cloudDb, user.uid, sesionActualizada);
      } else {
        guardarJsonLocal(LS_KEYS.config, nuevaConfig);
        guardarJsonLocal(LS_KEYS.sesion, sesionActualizada);
      }
      setConfigEquipo(nuevaConfig);
      setSesion(sesionActualizada);
      setPinAdminModal('');
      setPinVentasModal('');
      notify('Claves del equipo actualizadas.', 'green');
    } catch {
      notify('No se pudieron actualizar las claves.', 'red');
    } finally {
      setGuardandoAcceso(false);
    }
  };

  const handleReiniciarSistema = async (): Promise<void> => {
    if (!esAdmin) return;
    if (!confirmReiniciar) {
      setConfirmReiniciar(true);
      window.setTimeout(() => setConfirmReiniciar(false), 4000);
      return;
    }

    const idsABorrar = pedidos.map((p) => p.id);
    if (pendingUndoDelete) {
      window.clearTimeout(pendingUndoDelete.timerId);
      pendingDeleteIdsRef.current.clear();
      setPendingUndoDelete(null);
    }
    limpiarSeleccion();
    setPedidos([]);
    setDrawerPedidoId(null);
    setLineasBorrador([]);
    setAdvertenciasExcel([]);
    setCantidadesManuales({});
    setNombreArchivoCargado('');
    setForm(crearFormInicial());
    setConfirmReiniciar(false);
    setVista('pedidos');
    guardarJsonLocal(LS_KEYS.pedidos, []);

    if (cloudDb && user && idsABorrar.length > 0) {
      try {
        await clearAllOrdersInCloud(cloudDb, idsABorrar);
      } catch {
        notify('Reiniciado localmente; error al limpiar nube.', 'amber');
        return;
      }
    }
    notify('Sistema reiniciado en cero.', 'neutral');
  };

  // ============================================================================
  // PANTALLA DE INICIO DE SESIÓN / CONFIGURACIÓN INICIAL
  // ============================================================================

  if (!sesionValida) {
    const esPrimeraVez = !configEquipo;
    return (
      <div
        className="min-h-screen bg-[#FAFAFA] text-[#111111] flex flex-col items-center justify-center px-4 py-12 antialiased selection:bg-[#FCE7F3]"
        style={{ fontFamily: "'Plus Jakarta Sans', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif" }}
      >
        <div className="w-full max-w-sm bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] p-7 space-y-6">
          <div className="flex flex-col items-center text-center">
            <div className="pt-2 pb-1 flex flex-col items-center justify-center">
              <MibiLogo width={108} variant="pink" />
              <span className="mt-2 inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-medium bg-zinc-50 text-zinc-600 border border-zinc-200/80 tracking-tight">
                Gestión Mayorista
              </span>
            </div>

            <h1 className="mt-5 text-base font-semibold tracking-tight text-[#111111]">
              {esPrimeraVez ? 'Configurar acceso del equipo' : 'Ingreso al equipo'}
            </h1>
            <p className="mt-1 text-xs text-zinc-500">
              {esPrimeraVez
                ? 'Creá las claves para proteger pagos, cheques y borrados antes de compartir el enlace.'
                : 'Ingresá tu nombre y clave para continuar.'}
            </p>
          </div>

          {esPrimeraVez ? (
            <form onSubmit={handleConfigurarAccesoInicial} className="space-y-4">
              <div>
                <label className="block text-xs font-medium text-zinc-600 mb-1.5">Tu nombre</label>
                <input
                  type="text"
                  value={nombreIngreso}
                  onChange={(e) => setNombreIngreso(e.target.value)}
                  placeholder="Ej. Rocío"
                  autoFocus
                  className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 focus:ring-2 focus:ring-pink-200/60 transition-all"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-zinc-600 mb-1.5">
                  Clave Administración <span className="text-zinc-400 font-normal">(control total)</span>
                </label>
                <input
                  type="password"
                  value={pinAdminNuevo}
                  onChange={(e) => setPinAdminNuevo(e.target.value)}
                  placeholder="Mínimo 4 dígitos o letras"
                  className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 focus:ring-2 focus:ring-pink-200/60 transition-all tabular-nums"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-zinc-600 mb-1.5">
                  Clave Ventas / Depósito <span className="text-zinc-400 font-normal">(opcional, sin tocar pagos)</span>
                </label>
                <input
                  type="password"
                  value={pinVentasNuevo}
                  onChange={(e) => setPinVentasNuevo(e.target.value)}
                  placeholder="Distinta a Administración"
                  className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 focus:ring-2 focus:ring-pink-200/60 transition-all tabular-nums"
                />
              </div>

              {errorIngreso && <p className="text-xs text-red-600">{errorIngreso}</p>}

              <button
                type="submit"
                disabled={guardandoAcceso}
                className="w-full py-2.5 rounded-full text-sm font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors disabled:opacity-50"
              >
                {guardandoAcceso ? 'Activando...' : 'Activar acceso y entrar'}
              </button>
            </form>
          ) : (
            <form onSubmit={handleIngresarEquipo} className="space-y-4">
              <div>
                <label className="block text-xs font-medium text-zinc-600 mb-1.5">Tu nombre</label>
                <input
                  type="text"
                  value={nombreIngreso}
                  onChange={(e) => setNombreIngreso(e.target.value)}
                  placeholder="Ej. Juancho, Rocío, Depósito..."
                  autoFocus
                  className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 focus:ring-2 focus:ring-pink-200/60 transition-all"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-zinc-600 mb-1.5">Clave del equipo</label>
                <input
                  type="password"
                  value={pinIngreso}
                  onChange={(e) => setPinIngreso(e.target.value)}
                  placeholder="Ingresá tu clave"
                  className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 focus:ring-2 focus:ring-pink-200/60 transition-all tabular-nums"
                />
              </div>

              {errorIngreso && <p className="text-xs text-red-600">{errorIngreso}</p>}

              <button
                type="submit"
                disabled={guardandoAcceso}
                className="w-full py-2.5 rounded-full text-sm font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors disabled:opacity-50"
              >
                {guardandoAcceso ? 'Verificando...' : 'Ingresar'}
              </button>
            </form>
          )}
        </div>

        <p className="mt-6 text-[11px] text-zinc-400 text-center">{MIBI_BRAND.pieOficial}</p>
      </div>
    );
  }

  // ============================================================================
  // INTERFAZ PRINCIPAL AUTENTICADA (2 VISTAS + DRAWER + BARRA FLOTANTE)
  // ============================================================================

  const catalogoFiltradoManual = CATALOGO_MIBI.filter((prod) => {
    const q = normalizarTexto(busquedaCatalogo);
    if (!q) return true;
    return (
      normalizarTexto(prod.nombre).includes(q) ||
      normalizarTexto(prod.rubro).includes(q) ||
      prod.variantes.some((v) => v.codigo.includes(q))
    );
  });

  return (
    <div
      className="min-h-screen bg-[#FAFAFA] text-[#111111] flex flex-col antialiased selection:bg-[#FCE7F3]"
      style={{ fontFamily: "'Plus Jakarta Sans', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif" }}
    >
      <style>{`
        @media print {
          body * { visibility: hidden; }
          .mibi-print-area, .mibi-print-area * { visibility: visible; }
          .mibi-print-area { position: absolute; left: 0; top: 0; width: 100%; background: #ffffff; padding: 24px; }
        }
      `}</style>

      {/* INPUTS OCULTOS PARA IMPORTACIÓN EXCEL */}
      <input
        ref={fileInputListaRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) void handleArchivoLista006(file);
          e.target.value = '';
        }}
      />
      <input
        ref={fileInputGeneralRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) void handleImportarPlanillaGeneral(file);
          e.target.value = '';
        }}
      />

      {/* TOAST SILENCIOSO */}
      {toast && (
        <div
          className={`fixed left-1/2 -translate-x-1/2 z-50 pointer-events-none transition-all ${
            (vista === 'pedidos' && selectedIds.size > 0) || pendingUndoDelete ? 'bottom-24' : 'bottom-6'
          }`}
        >
          <div className="flex items-center gap-2 px-4 py-2.5 rounded-full bg-[#111111] text-white text-xs font-medium shadow-[0_4px_16px_rgba(0,0,0,0.12)]">
            <span
              className={`w-1.5 h-1.5 rounded-full ${
                toast.tone === 'green'
                  ? 'bg-emerald-400'
                  : toast.tone === 'amber'
                  ? 'bg-amber-400'
                  : toast.tone === 'red'
                  ? 'bg-red-400'
                  : 'bg-[#F4B6CE]'
              }`}
            />
            <span>{toast.texto}</span>
          </div>
        </div>
      )}

      {/* TOAST INTERACTIVO DE DESHACER ELIMINACIÓN (6 SEGUNDOS) */}
      {pendingUndoDelete && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 z-50">
          <div className="flex items-center gap-3 pl-4 pr-2 py-2 rounded-full bg-zinc-900 text-white text-xs font-medium shadow-lg border border-zinc-800">
            <span>
              {pendingUndoDelete.pedidos.length === 1
                ? `Pedido #${pendingUndoDelete.pedidos[0].numeroPedido} eliminado`
                : `${pendingUndoDelete.pedidos.length} pedidos eliminados`}
            </span>
            <button
              type="button"
              onClick={handleDeshacerEliminacion}
              className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-white text-[#111111] hover:bg-[#FCE7F3] font-semibold transition-colors"
            >
              <Undo2 className="w-3.5 h-3.5" strokeWidth={1.75} />
              Deshacer
            </button>
          </div>
        </div>
      )}

      {/* LÍNEA SUPERIOR ROSA MIBI® (2PX) + BARRA STICKY 60PX */}
      <div className="h-[2px] w-full bg-[#F4B6CE]" />
      <header className="sticky top-0 z-30 h-[60px] bg-white/95 backdrop-blur border-b border-[#EAEAEA]">
        <div className="max-w-7xl mx-auto h-full px-4 sm:px-6 flex items-center justify-between gap-3">
          {/* Izquierda: Logo + estado sincronización */}
          <div className="flex items-center gap-3">
            <div className="inline-flex items-center">
              <MibiLogo width={64} variant="pink" />
            </div>
            <span className="hidden md:inline-flex items-center gap-1.5 text-[11px] text-zinc-400">
              <span
                className={`w-1.5 h-1.5 rounded-full ${
                  errorNube ? 'bg-red-500' : cloudDb ? 'bg-emerald-500' : 'bg-zinc-400'
                }`}
              />
              {errorNube ? 'Sin conexión nube' : cloudDb ? 'Nube en vivo' : 'Modo local'}
              {pedidosInvalidosNube > 0 ? ` (${pedidosInvalidosNube} omitidos)` : ''}
            </span>
          </div>

          {/* Centro: Selector segmentado Pedidos / Nuevo pedido */}
          <nav className="inline-flex items-center p-1 rounded-full bg-zinc-100 border border-zinc-200/70">
            <button
              type="button"
              onClick={() => setVista('pedidos')}
              className={`px-4 py-1.5 rounded-full text-xs font-medium transition-all ${
                vista === 'pedidos'
                  ? 'bg-white text-[#111111] shadow-[0_1px_2px_rgba(0,0,0,0.06)]'
                  : 'text-zinc-500 hover:text-[#111111]'
              }`}
            >
              Pedidos
            </button>
            <button
              type="button"
              onClick={() => setVista('nuevo')}
              className={`inline-flex items-center gap-1.5 px-4 py-1.5 rounded-full text-xs font-medium transition-all ${
                vista === 'nuevo'
                  ? 'bg-white text-[#111111] shadow-[0_1px_2px_rgba(0,0,0,0.06)]'
                  : 'text-zinc-500 hover:text-[#111111]'
              }`}
            >
              <span>Nuevo pedido</span>
              {hayBorradorSinGuardar && <span className="w-1.5 h-1.5 rounded-full bg-amber-500" />}
            </button>
          </nav>

          {/* Derecha: Acciones y usuario */}
          <div className="flex items-center gap-2">
            {pedidos.length > 0 && (
              <button
                type="button"
                onClick={() => {
                  if (!XLSX) {
                    notify('Preparando Excel, intentá en un segundo.');
                    return;
                  }
                  void exportarTableroMaestro(pedidos, resumenFinanciero, XLSX).then(() => {
                    notify(`Exportado PEDIDOS_MAYORISTA_MIBI (${pedidos.length} pedidos).`, 'green');
                  });
                }}
                className="hidden sm:inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors"
              >
                <Download className="w-3.5 h-3.5" strokeWidth={1.5} />
                Exportar Excel
              </button>
            )}

            <button
              type="button"
              onClick={() => fileInputGeneralRef.current?.click()}
              className="hidden lg:inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors"
              title="Importar planilla general previa PEDIDOS MAYORISTA MIBI.xlsx"
            >
              <Upload className="w-3.5 h-3.5" strokeWidth={1.5} />
              Importar planilla
            </button>

            {esAdmin && pedidos.length > 0 && (
              <button
                type="button"
                onClick={() => void handleReiniciarSistema()}
                className={`hidden sm:inline-flex items-center gap-1 px-3 py-1.5 rounded-full text-xs font-medium border transition-colors ${
                  confirmReiniciar
                    ? 'bg-red-600 text-white border-red-600'
                    : 'bg-white text-zinc-500 border-zinc-200 hover:text-red-700 hover:border-red-200'
                }`}
              >
                <RotateCcw className="w-3.5 h-3.5" strokeWidth={1.5} />
                {confirmReiniciar ? 'Confirmar reinicio' : 'Reiniciar'}
              </button>
            )}

            <button
              type="button"
              onClick={() => setModalEquipoAbierto(true)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-zinc-50 border border-zinc-200/80 text-zinc-700 hover:bg-zinc-100 transition-colors"
              title="Equipo, claves y verificación QA"
            >
              <Users className="w-3.5 h-3.5 text-zinc-500" strokeWidth={1.5} />
              <span className="hidden sm:inline">{sesion?.nombre}</span>
            </button>

            <button
              type="button"
              onClick={() => void handleCerrarSesion()}
              className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-800 hover:bg-zinc-100 transition-colors"
              title="Bloquear sesión"
              aria-label="Bloquear sesión"
            >
              <LogOut className="w-4 h-4" strokeWidth={1.5} />
            </button>
          </div>
        </div>
      </header>

      {/* CUERPO PRINCIPAL */}
      <main className="flex-1 w-full max-w-7xl mx-auto px-4 sm:px-6 py-7">
        {/* ================================================================ */}
        {/* VISTA 1: PEDIDOS                                                 */}
        {/* ================================================================ */}
        {vista === 'pedidos' && (
          <div className="space-y-6">
            {/* Franja superior de 4 indicadores limpios */}
            <section className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-5 py-4">
                <div className="text-xs font-medium text-zinc-500">Por armar</div>
                <div className="mt-1.5 text-2xl font-semibold tracking-tight text-[#111111] tabular-nums">
                  {resumenFinanciero.porArmar}
                </div>
              </div>

              <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-5 py-4">
                <div className="text-xs font-medium text-zinc-500">Cobrado en mano</div>
                <div className="mt-1.5 text-2xl font-semibold tracking-tight text-emerald-700 tabular-nums">
                  {formatARSCompacto(resumenFinanciero.cobradoEnMano)}
                </div>
              </div>

              <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-5 py-4">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs font-medium text-zinc-500">En cheques 30d</span>
                  {resumenFinanciero.chequesPorVencer > 0 && (
                    <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-red-50 text-red-700 border border-red-200/70 tabular-nums">
                      {resumenFinanciero.chequesPorVencer} al cobro
                    </span>
                  )}
                </div>
                <div className="mt-1.5 text-2xl font-semibold tracking-tight text-amber-700 tabular-nums">
                  {formatARSCompacto(resumenFinanciero.enCheques)}
                </div>
              </div>

              <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-5 py-4">
                <div className="text-xs font-medium text-zinc-500">Falta pago</div>
                <div className="mt-1.5 text-2xl font-semibold tracking-tight text-[#111111] tabular-nums">
                  {formatARSCompacto(resumenFinanciero.faltaPago)}
                </div>
              </div>
            </section>

            {/* Estado vacío inicial (arranque 100% en cero) */}
            {cargandoNube ? (
              <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] py-16 text-center text-sm text-zinc-400">
                Sincronizando pedidos...
              </div>
            ) : pedidos.length === 0 ? (
              <section className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-6 py-16 text-center space-y-4">
                <div className="mx-auto w-11 h-11 rounded-full bg-zinc-100 flex items-center justify-center text-zinc-500">
                  <Package className="w-5 h-5" strokeWidth={1.5} />
                </div>
                <div className="space-y-1">
                  <h2 className="text-base font-semibold text-[#111111]">Todavía no hay pedidos cargados</h2>
                  <p className="text-xs text-zinc-500 max-w-md mx-auto">
                    Subí la planilla oficial <strong className="font-medium text-zinc-700">LISTA N° 006</strong> de un
                    cliente o importá tu planilla general para empezar a operar.
                  </p>
                </div>
                <div className="flex flex-wrap items-center justify-center gap-2.5 pt-2">
                  <button
                    type="button"
                    onClick={() => setVista('nuevo')}
                    className="inline-flex items-center gap-1.5 px-5 py-2.5 rounded-full text-xs font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors"
                  >
                    <Plus className="w-3.5 h-3.5" strokeWidth={1.75} />
                    Nuevo pedido
                  </button>
                  <button
                    type="button"
                    onClick={() => fileInputGeneralRef.current?.click()}
                    className="inline-flex items-center gap-1.5 px-4 py-2.5 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors"
                  >
                    <Upload className="w-3.5 h-3.5" strokeWidth={1.5} />
                    Importar planilla previa
                  </button>
                </div>
              </section>
            ) : (
              <section className="space-y-4">
                {/* Buscador + Filtros rápidos */}
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                  <div className="relative w-full md:w-80">
                    <Search
                      className="w-4 h-4 text-zinc-400 absolute left-3.5 top-1/2 -translate-y-1/2"
                      strokeWidth={1.5}
                    />
                    <input
                      ref={searchInputRef}
                      type="text"
                      value={busqueda}
                      onChange={(e) => setBusqueda(e.target.value)}
                      placeholder="Buscar cliente o vendedor..."
                      className="w-full pl-9 pr-10 py-2 text-sm bg-white border border-zinc-200/90 rounded-full focus:outline-none focus:border-zinc-400 transition-colors placeholder:text-zinc-400"
                    />
                    {busqueda ? (
                      <button
                        type="button"
                        onClick={() => setBusqueda('')}
                        className="absolute right-3 top-1/2 -translate-y-1/2 text-zinc-400 hover:text-zinc-700"
                        aria-label="Limpiar búsqueda"
                      >
                        <X className="w-3.5 h-3.5" strokeWidth={1.5} />
                      </button>
                    ) : (
                      <kbd className="hidden sm:inline-block absolute right-3 top-1/2 -translate-y-1/2 text-[10px] font-medium text-zinc-400 bg-zinc-100 border border-zinc-200/80 rounded px-1.5 py-0.5 pointer-events-none">
                        /
                      </kbd>
                    )}
                  </div>

                  <div className="flex flex-wrap items-center gap-1.5">
                    {(
                      [
                        { id: 'todos', label: 'Todos', count: conteosFiltro.todos },
                        { id: 'falta_talle', label: 'Falta talle', count: conteosFiltro.falta_talle },
                        { id: 'cheques', label: 'Cheques', count: conteosFiltro.cheques },
                        { id: 'falta_pago', label: 'Falta pago', count: conteosFiltro.falta_pago },
                        { id: 'listos', label: 'Listos', count: conteosFiltro.listos }
                      ] as const
                    ).map((item) => {
                      const activo = filtro === item.id;
                      return (
                        <button
                          key={item.id}
                          type="button"
                          onClick={() => setFiltro(item.id)}
                          className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium border transition-colors ${
                            activo
                              ? 'bg-zinc-900 text-white border-zinc-900'
                              : 'bg-white text-zinc-600 border-zinc-200/90 hover:bg-zinc-50'
                          }`}
                        >
                          <span>{item.label}</span>
                          <span
                            className={`text-[11px] tabular-nums ${
                              activo ? 'text-zinc-300' : 'text-zinc-400'
                            }`}
                          >
                            {item.count}
                          </span>
                        </button>
                      );
                    })}

                    <button
                      type="button"
                      onClick={() => setVista('nuevo')}
                      className="ml-1 inline-flex items-center gap-1.5 px-4 py-1.5 rounded-full text-xs font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors"
                    >
                      <Plus className="w-3.5 h-3.5" strokeWidth={1.75} />
                      Nuevo pedido
                    </button>
                  </div>
                </div>

                {/* Tabla de 6 columnas esenciales + columna minimalista de selección múltiple */}
                <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b border-zinc-200/70 text-[11px] font-medium text-zinc-400">
                          <th className="py-3.5 pl-5 pr-2 w-10">
                            <input
                              ref={headerCheckboxRef}
                              type="checkbox"
                              checked={todosVisiblesSeleccionados}
                              onChange={handleToggleSelectAllVisible}
                              disabled={pedidosFiltrados.length === 0}
                              aria-label="Seleccionar todos los pedidos visibles"
                              className="w-4 h-4 rounded border-zinc-300 text-zinc-900 accent-zinc-900 cursor-pointer align-middle disabled:opacity-30"
                            />
                          </th>
                          <th className="py-3.5 pl-2 pr-3">N.º y Fecha</th>
                          <th className="py-3.5 px-3">Cliente</th>
                          <th className="py-3.5 px-3 text-right">Total</th>
                          <th className="py-3.5 px-3">Pago</th>
                          <th className="py-3.5 px-3">Armado</th>
                          <th className="py-3.5 pl-3 pr-6 text-right">Acción</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-zinc-200/60 text-sm">
                        {pedidosFiltrados.length === 0 ? (
                          <tr>
                            <td colSpan={7} className="py-14 text-center text-sm text-zinc-400">
                              No hay pedidos con ese filtro.
                            </td>
                          </tr>
                        ) : (
                          pedidosFiltrados.map((p) => {
                            const faltaDefinir = p.faltantesVariante > 0;
                            const esChequePendiente = p.medioPago === 'Cheque' && !p.chequeAcreditado;
                            const diasCheque = getDiasParaAcreditacionCheque(p);
                            const chequeAlCobro = esChequePendiente && diasCheque !== null && diasCheque <= 0;
                            const atrasado = esDespachoAtrasado(p);
                            const seleccionado = selectedIds.has(p.id);

                            return (
                              <tr
                                key={p.id}
                                onClick={() => setDrawerPedidoId(p.id)}
                                className={`cursor-pointer transition-colors ${
                                  seleccionado ? 'bg-pink-50/45 hover:bg-pink-50/70' : 'hover:bg-zinc-50/80'
                                }`}
                              >
                                <td
                                  className="py-3.5 pl-5 pr-2 w-10"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleToggleSelectRow(p.id, e.shiftKey);
                                  }}
                                >
                                  <input
                                    type="checkbox"
                                    checked={seleccionado}
                                    onChange={() => {
                                      // Controlado en onClick de la celda para capturar Shift + Clic
                                    }}
                                    aria-label={`Seleccionar pedido #${p.numeroPedido}`}
                                    className="w-4 h-4 rounded border-zinc-300 text-zinc-900 accent-zinc-900 cursor-pointer align-middle pointer-events-none"
                                  />
                                </td>

                                <td className="py-3.5 pl-2 pr-3 whitespace-nowrap tabular-nums">
                                  <div className="font-semibold text-[#111111]">#{p.numeroPedido}</div>
                                  <div className="text-xs text-zinc-400 mt-0.5 flex items-center gap-1.5">
                                    <span>{formatShortDate(p.fechaPedido)}</span>
                                    {atrasado && (
                                      <span
                                        className="w-1.5 h-1.5 rounded-full bg-amber-500"
                                        title={`Despacho previsto: ${formatShortDate(p.despachoPrevisto)}`}
                                      />
                                    )}
                                  </div>
                                </td>

                                <td className="py-3.5 px-3 min-w-[200px]">
                                  <div className="font-medium text-[#111111]">{p.cliente}</div>
                                  <div className="text-xs text-zinc-400 mt-0.5">
                                    {p.vendedor}
                                    {p.detalleResumen && p.detalleResumen !== '—' ? ` · ${p.detalleResumen}` : ''}
                                  </div>
                                </td>

                                <td className="py-3.5 px-3 text-right whitespace-nowrap tabular-nums">
                                  <div className="font-semibold text-[#111111]">{formatARS(p.totalFacturado)}</div>
                                  <div className="text-xs text-zinc-400 mt-0.5">
                                    {esChequePendiente
                                      ? chequeAlCobro
                                        ? diasCheque === 0
                                          ? 'Vence hoy'
                                          : `Vencido hace ${Math.abs(diasCheque ?? 0)}d`
                                        : `Faltan ${diasCheque}d · ${formatShortDate(p.fechaAcreditacionCheque)}`
                                      : p.saldoPendiente > 0.01
                                      ? `Resta ${formatARSCompacto(p.saldoPendiente)}`
                                      : 'Cobrado'}
                                  </div>
                                </td>

                                <td className="py-3.5 px-3 whitespace-nowrap" onClick={(e) => e.stopPropagation()}>
                                  <div className="flex flex-col items-start gap-1.5">
                                    {p.estadoPago === 'Pagado' && <StatusPill tone="green">Pagado</StatusPill>}
                                    {p.estadoPago === 'Pendiente' && (
                                      <StatusPill tone={chequeAlCobro ? 'red' : 'amber'}>
                                        {esChequePendiente
                                          ? chequeAlCobro
                                            ? 'Cheque al cobro'
                                            : 'Pendiente · Cheque'
                                          : 'Pendiente'}
                                      </StatusPill>
                                    )}
                                    {p.estadoPago === 'Falta pago' && <StatusPill tone="red">Falta pago</StatusPill>}

                                    {esChequePendiente && esAdmin && (
                                      <button
                                        type="button"
                                        onClick={() => confirmarAcreditacion(p)}
                                        className="text-[11px] font-medium text-emerald-700 hover:text-emerald-800 underline underline-offset-2"
                                      >
                                        Confirmar plata acreditada
                                      </button>
                                    )}
                                  </div>
                                </td>

                                <td className="py-3.5 px-3 whitespace-nowrap">
                                  <div className="flex flex-col items-start gap-0.5">
                                    {faltaDefinir ? (
                                      <StatusPill tone="red">Falta talle o color</StatusPill>
                                    ) : p.estadoArmado === 'Listo' ? (
                                      <StatusPill tone="green">Listo</StatusPill>
                                    ) : p.estadoArmado === 'En armado' ? (
                                      <StatusPill tone="ink">En armado</StatusPill>
                                    ) : (
                                      <StatusPill tone="neutral">Por armar</StatusPill>
                                    )}
                                    {p.estadoEnvio !== 'Pendiente' ? (
                                      <span className="text-[11px] text-zinc-400 pl-0.5">
                                        {p.estadoEnvio}
                                        {p.transporte ? ` · ${p.transporte}` : ''}
                                      </span>
                                    ) : atrasado ? (
                                      <span className="text-[11px] text-amber-700 pl-0.5">Despacho atrasado</span>
                                    ) : null}
                                  </div>
                                </td>

                                <td
                                  className="py-3.5 pl-3 pr-6 text-right whitespace-nowrap"
                                  onClick={(e) => e.stopPropagation()}
                                >
                                  <div className="inline-flex items-center gap-1.5">
                                    {faltaDefinir ? (
                                      <button
                                        type="button"
                                        onClick={() => setDrawerPedidoId(p.id)}
                                        className="px-3 py-1.5 rounded-full text-xs font-medium bg-red-50 text-red-700 border border-red-200/70 hover:bg-red-100/70 transition-colors"
                                      >
                                        Completar
                                      </button>
                                    ) : (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const res = exportarHojaDeposito(p, XLSX);
                                          if (!res.success) {
                                            notify(res.errors[0] ?? 'No se pudo exportar este pedido.', 'red');
                                            return;
                                          }
                                          notify(`Descargada Hoja depósito #${p.numeroPedido}`, 'green');
                                        }}
                                        className="px-3 py-1.5 rounded-full text-xs font-medium bg-white text-zinc-700 border border-zinc-200 hover:bg-zinc-50 transition-colors"
                                      >
                                        Hoja depósito
                                      </button>
                                    )}
                                    <button
                                      type="button"
                                      onClick={() => setDrawerPedidoId(p.id)}
                                      className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-800 hover:bg-zinc-100 transition-colors"
                                      aria-label="Ver detalle"
                                    >
                                      <ChevronRight className="w-4 h-4" strokeWidth={1.5} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </section>
            )}
          </div>
        )}

        {/* ================================================================ */}
        {/* VISTA 2: NUEVO PEDIDO (max-w-2xl mx-auto)                        */}
        {/* ================================================================ */}
        {vista === 'nuevo' && (
          <div className="max-w-2xl mx-auto space-y-6">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <button
                  type="button"
                  onClick={() => setVista('pedidos')}
                  className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-800 hover:bg-zinc-100 transition-colors"
                  aria-label="Volver a pedidos"
                >
                  <ArrowLeft className="w-4 h-4" strokeWidth={1.5} />
                </button>
                <div>
                  <h1 className="text-lg font-semibold tracking-tight text-[#111111]">
                    Nuevo pedido <span className="text-zinc-400 font-normal tabular-nums">#{proximoNumeroPedido}</span>
                  </h1>
                  <p className="text-xs text-zinc-500">Cargá el Excel Lista N° 006 o ingresá cantidades a mano.</p>
                </div>
              </div>
            </div>

            {/* PASO 1: DATOS DEL PEDIDO */}
            <section className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] p-6 space-y-4">
              <h2 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">1. Datos del pedido</h2>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="sm:col-span-2">
                  <label className="block text-xs font-medium text-zinc-600 mb-1.5">Cliente o negocio *</label>
                  <input
                    type="text"
                    value={form.cliente}
                    onChange={(e) => setForm((prev) => ({ ...prev, cliente: e.target.value }))}
                    placeholder="Ej. Casa Lorenzo, Pañalera Nuevo Mundo..."
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-zinc-600 mb-1.5">Vendedor</label>
                  <select
                    value={form.vendedor}
                    onChange={(e) => setForm((prev) => ({ ...prev, vendedor: e.target.value }))}
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900"
                  >
                    {VENDEDORES_MIBI.map((v) => (
                      <option key={v} value={v}>
                        {v}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-medium text-zinc-600 mb-1.5">WhatsApp del cliente</label>
                  <input
                    type="text"
                    value={form.whatsapp}
                    onChange={(e) => setForm((prev) => ({ ...prev, whatsapp: e.target.value }))}
                    placeholder="Ej. 11-5606-9384"
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 tabular-nums"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-zinc-600 mb-1.5">Fecha pedido</label>
                  <input
                    type="date"
                    value={form.fechaPedido}
                    onChange={(e) => setForm((prev) => ({ ...prev, fechaPedido: e.target.value }))}
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 tabular-nums"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-zinc-600 mb-1.5">Despacho previsto</label>
                  <input
                    type="date"
                    value={form.despachoPrevisto}
                    onChange={(e) => setForm((prev) => ({ ...prev, despachoPrevisto: e.target.value }))}
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 tabular-nums"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-zinc-600 mb-1.5">Descuento (%)</label>
                  <input
                    type="number"
                    min={0}
                    max={100}
                    step="0.5"
                    value={form.descuentoPorc}
                    onChange={(e) =>
                      setForm((prev) => ({
                        ...prev,
                        descuentoPorc: Math.max(0, Math.min(100, Number(e.target.value) || 0))
                      }))
                    }
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 tabular-nums"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-zinc-600 mb-1.5">Observaciones</label>
                  <input
                    type="text"
                    value={form.observaciones}
                    onChange={(e) => setForm((prev) => ({ ...prev, observaciones: e.target.value }))}
                    placeholder="Notas internas o entrega..."
                    className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900"
                  />
                </div>
              </div>

              {/* Condición inicial de pago (restringida para Ventas/Depósito) */}
              <div className="pt-3 border-t border-zinc-200/70 space-y-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-medium text-zinc-600">Estado de pago inicial</span>
                  {!esAdmin && (
                    <span className="text-[11px] text-zinc-400 inline-flex items-center gap-1">
                      <Lock className="w-3 h-3" strokeWidth={1.5} /> Solo Administración confirma pagos
                    </span>
                  )}
                </div>

                <div className="grid grid-cols-3 gap-2">
                  {(['Sin pagar', 'Transferencia / Efectivo', 'Cheque'] as const).map((tipo) => {
                    const bloqueado = !esAdmin && tipo !== 'Sin pagar';
                    const activo = form.pagoTipo === tipo;
                    return (
                      <button
                        key={tipo}
                        type="button"
                        disabled={bloqueado}
                        onClick={() => setForm((prev) => ({ ...prev, pagoTipo: tipo }))}
                        className={`px-3 py-2 rounded-xl text-xs font-medium border transition-colors disabled:opacity-40 disabled:cursor-not-allowed ${
                          activo
                            ? 'bg-zinc-900 text-white border-zinc-900'
                            : 'bg-white text-zinc-700 border-zinc-200 hover:bg-zinc-50'
                        }`}
                      >
                        {tipo === 'Transferencia / Efectivo' ? 'Pagado / Seña' : tipo}
                      </button>
                    );
                  })}
                </div>

                {esAdmin && form.pagoTipo === 'Transferencia / Efectivo' && (
                  <div>
                    <label className="block text-xs font-medium text-zinc-600 mb-1">
                      Monto acreditado (vacío = total facturado)
                    </label>
                    <input
                      type="text"
                      value={form.montoAbonadoInput}
                      onChange={(e) => setForm((prev) => ({ ...prev, montoAbonadoInput: e.target.value }))}
                      placeholder={totalesBorrador.totalFacturado > 0 ? String(totalesBorrador.totalFacturado) : 'Total'}
                      className="w-full px-3.5 py-2 text-sm bg-white border border-zinc-200 rounded-xl tabular-nums"
                    />
                  </div>
                )}

                {esAdmin && form.pagoTipo === 'Cheque' && (
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                      <label className="block text-xs font-medium text-zinc-600 mb-1">Fecha recepción</label>
                      <input
                        type="date"
                        value={form.fechaRecepcionCheque}
                        onChange={(e) =>
                          setForm((prev) => ({
                            ...prev,
                            fechaRecepcionCheque: e.target.value,
                            fechaAcreditacionCheque: addDaysISO(e.target.value, 30)
                          }))
                        }
                        className="w-full px-3 py-2 text-xs bg-white border border-zinc-200 rounded-xl tabular-nums"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-zinc-600 mb-1">Acreditación (+30d)</label>
                      <input
                        type="date"
                        value={form.fechaAcreditacionCheque}
                        onChange={(e) => setForm((prev) => ({ ...prev, fechaAcreditacionCheque: e.target.value }))}
                        className="w-full px-3 py-2 text-xs bg-white border border-zinc-200 rounded-xl tabular-nums"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-zinc-600 mb-1">Detalle cheque</label>
                      <input
                        type="text"
                        value={form.detalleCheque}
                        onChange={(e) => setForm((prev) => ({ ...prev, detalleCheque: e.target.value }))}
                        placeholder="Banco / N.º"
                        className="w-full px-3 py-2 text-xs bg-white border border-zinc-200 rounded-xl"
                      />
                    </div>
                  </div>
                )}
              </div>
            </section>

            {/* PASO 2: SUBIR EXCEL LISTA N° 006 O CARGA MANUAL */}
            <section className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] p-6 space-y-4">
              <div className="flex items-center justify-between gap-2">
                <h2 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">
                  2. Artículos (LISTA N° 006)
                </h2>
                <div className="inline-flex items-center p-0.5 rounded-full bg-zinc-100 border border-zinc-200/70">
                  <button
                    type="button"
                    onClick={() => setModoCargaManual(false)}
                    className={`px-3 py-1 rounded-full text-xs font-medium transition-colors ${
                      !modoCargaManual ? 'bg-white text-[#111111] shadow-sm' : 'text-zinc-500'
                    }`}
                  >
                    Subir Excel
                  </button>
                  <button
                    type="button"
                    onClick={() => setModoCargaManual(true)}
                    className={`px-3 py-1 rounded-full text-xs font-medium transition-colors ${
                      modoCargaManual ? 'bg-white text-[#111111] shadow-sm' : 'text-zinc-500'
                    }`}
                  >
                    Cargar cantidades a mano
                  </button>
                </div>
              </div>

              {!modoCargaManual ? (
                <div
                  onDragOver={(e) => {
                    e.preventDefault();
                    setArrastrandoArchivo(true);
                  }}
                  onDragLeave={() => setArrastrandoArchivo(false)}
                  onDrop={(e) => {
                    e.preventDefault();
                    setArrastrandoArchivo(false);
                    const file = e.dataTransfer.files?.[0];
                    if (file) void handleArchivoLista006(file);
                  }}
                  onClick={() => fileInputListaRef.current?.click()}
                  className={`cursor-pointer rounded-2xl border-2 border-dashed px-6 py-8 text-center transition-colors ${
                    arrastrandoArchivo
                      ? 'border-[#F4B6CE] bg-pink-50/40'
                      : 'border-zinc-200 hover:border-zinc-300 bg-[#FAFAFA]'
                  }`}
                >
                  <FileSpreadsheet className="w-7 h-7 mx-auto text-zinc-400" strokeWidth={1.5} />
                  <p className="mt-2.5 text-sm font-medium text-[#111111]">
                    {nombreArchivoCargado
                      ? `Cargado: ${nombreArchivoCargado}`
                      : 'Arrastrá la planilla LISTA N° 006 o hacé clic para elegir'}
                  </p>
                  <p className="mt-1 text-xs text-zinc-400">Soporta .xlsx, .xls y .csv · Lectura determinística 100% local</p>
                </div>
              ) : (
                <div className="space-y-3">
                  <input
                    type="text"
                    value={busquedaCatalogo}
                    onChange={(e) => setBusquedaCatalogo(e.target.value)}
                    placeholder="Filtrar catálogo por nombre, código o rubro..."
                    className="w-full px-3.5 py-2 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900"
                  />
                  <div className="max-h-72 overflow-y-auto divide-y divide-zinc-200/60 border border-zinc-200/80 rounded-xl">
                    {catalogoFiltradoManual.map((prod) => {
                      const val = cantidadesManuales[prod.idGrupo] ?? 0;
                      return (
                        <div key={prod.idGrupo} className="px-3.5 py-2.5 flex items-center justify-between gap-3">
                          <div className="min-w-0">
                            <div className="text-xs font-medium text-[#111111] truncate">{prod.nombre}</div>
                            <div className="text-[11px] text-zinc-400 tabular-nums">
                              {prod.variantes.map((v) => v.codigo).join(' / ')} · {prod.rubro} ·{' '}
                              {formatARSCompacto(prod.precioNeto)} neto
                            </div>
                          </div>
                          <input
                            type="number"
                            min={0}
                            value={val === 0 ? '' : val}
                            onChange={(e) => handleCambioCantidadManual(prod, e.target.value)}
                            placeholder="0"
                            className="w-20 px-2.5 py-1.5 text-sm text-right bg-white border border-zinc-200 rounded-lg tabular-nums focus:outline-none focus:border-zinc-900"
                          />
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              <div className="flex items-center justify-between pt-1">
                <button
                  type="button"
                  onClick={handleCargarEjemploLista006}
                  className="text-xs text-zinc-400 hover:text-[#111111] underline underline-offset-2 transition-colors"
                >
                  Probar con planilla de ejemplo
                </button>
                {lineasBorrador.length > 0 && (
                  <button
                    type="button"
                    onClick={() => {
                      setLineasBorrador([]);
                      setCantidadesManuales({});
                      setAdvertenciasExcel([]);
                      setNombreArchivoCargado('');
                    }}
                    className="text-xs text-zinc-400 hover:text-red-600 transition-colors"
                  >
                    Vaciar artículos
                  </button>
                )}
              </div>
            </section>

            {/* PASO 3: ALERTAS BLOQUEANTES, TALLES/COLORES, RESUMEN Y TOTALES */}
            {advertenciasExcel.length > 0 && (
              <section className="bg-red-50/70 border border-red-200 rounded-2xl p-5 space-y-3">
                <div className="flex items-center justify-between gap-2">
                  <div className="flex items-center gap-2 text-xs font-semibold text-red-800">
                    <AlertTriangle className="w-4 h-4 text-red-600" strokeWidth={1.75} />
                    <span>Errores bloqueantes detectados en el Excel ({advertenciasExcel.length})</span>
                  </div>
                  <button
                    type="button"
                    onClick={() => setAdvertenciasExcel([])}
                    className="text-[11px] font-medium text-red-700 hover:underline"
                  >
                    Confirmar y descartar avisos
                  </button>
                </div>
                <ul className="space-y-1 text-xs text-red-700">
                  {advertenciasExcel.map((adv, idx) => (
                    <li key={`${adv.hoja}-${adv.fila}-${idx}`} className="tabular-nums">
                      • Hoja <strong>{adv.hoja}</strong>, Fila <strong>{adv.fila}</strong> (cód. {adv.codigo}):{' '}
                      {adv.mensaje}
                    </li>
                  ))}
                </ul>
              </section>
            )}

            {totalesBorrador.faltantes > 0 && (
              <section className="bg-red-50/60 border border-red-200/80 rounded-2xl p-5 space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="text-xs font-semibold text-red-800">
                    Falta definir talle o color en {totalesBorrador.faltantes} artículo(s)
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      const msg = generarMensajeWhatsApp(form.cliente.trim(), proximoNumeroPedido, lineasBorrador);
                      copyTextSafe(msg);
                      setCopiadoId('borrador-wsp');
                      notify('Mensaje de consulta copiado para WhatsApp.', 'green');
                      window.setTimeout(() => setCopiadoId(null), 2200);
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-white border border-red-200 text-red-700 hover:bg-red-50"
                  >
                    {copiadoId === 'borrador-wsp' ? (
                      <Check className="w-3.5 h-3.5 text-emerald-600" strokeWidth={1.75} />
                    ) : (
                      <MessageCircle className="w-3.5 h-3.5" strokeWidth={1.5} />
                    )}
                    {copiadoId === 'borrador-wsp' ? 'WhatsApp copiado' : 'Copiar WhatsApp'}
                  </button>
                </div>

                <div className="space-y-2.5">
                  {lineasBorrador
                    .filter((l) => l.requiereRevision || l.estado !== 'OK')
                    .map((l) => {
                      const prod = CATALOGO_BY_GRUPO.get(l.idGrupo);
                      if (!prod || prod.variantes.length < 2) return null;
                      const [v0, v1] = prod.variantes;
                      const key = `borr_${l.idLinea}`;
                      const cur = resolucionesManual[key] ?? { q0: '0', q1: '0' };
                      return (
                        <div key={l.idLinea} className="bg-white rounded-xl border border-red-200/70 p-3.5 space-y-2">
                          <div className="flex items-center justify-between text-xs">
                            <span className="font-semibold text-[#111111]">{l.nombreProducto}</span>
                            <span className="font-semibold text-red-700 tabular-nums">Total: {l.cantidadTotal} u.</span>
                          </div>
                          {l.motivoRevision && <p className="text-[11px] text-red-600">{l.motivoRevision}</p>}
                          <div className="flex flex-wrap items-end gap-2">
                            <div className="flex-1 min-w-[110px]">
                              <label className="block text-[11px] text-zinc-500 mb-1">
                                {v0.etiqueta} ({v0.codigo})
                              </label>
                              <input
                                type="number"
                                min={0}
                                max={l.cantidadTotal}
                                value={cur.q0}
                                onChange={(e) =>
                                  setResolucionesManual((prev) => ({
                                    ...prev,
                                    [key]: { ...cur, q0: e.target.value }
                                  }))
                                }
                                className="w-full px-2.5 py-1.5 text-xs bg-white border border-zinc-200 rounded-lg tabular-nums"
                              />
                            </div>
                            <div className="flex-1 min-w-[110px]">
                              <label className="block text-[11px] text-zinc-500 mb-1">
                                {v1.etiqueta} ({v1.codigo})
                              </label>
                              <input
                                type="number"
                                min={0}
                                max={l.cantidadTotal}
                                value={cur.q1}
                                onChange={(e) =>
                                  setResolucionesManual((prev) => ({
                                    ...prev,
                                    [key]: { ...cur, q1: e.target.value }
                                  }))
                                }
                                className="w-full px-2.5 py-1.5 text-xs bg-white border border-zinc-200 rounded-lg tabular-nums"
                              />
                            </div>
                            <button
                              type="button"
                              onClick={() =>
                                handleResolverEnBorrador(
                                  l.idLinea,
                                  parseInt(cur.q0, 10) || 0,
                                  parseInt(cur.q1, 10) || 0
                                )
                              }
                              className="px-3.5 py-1.5 rounded-full text-xs font-medium bg-zinc-900 hover:bg-zinc-800 text-white"
                            >
                              Confirmar
                            </button>
                          </div>
                        </div>
                      );
                    })}
                </div>
              </section>
            )}

            {lineasBorrador.length > 0 && (
              <section className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] p-6 space-y-4">
                <div className="flex items-center justify-between">
                  <h2 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">
                    3. Resumen y liquidación ({lineasBorrador.length} ítems · {totalesBorrador.totalUnidades} u.)
                  </h2>
                </div>

                <div className="max-h-60 overflow-y-auto divide-y divide-zinc-200/60 border border-zinc-200/70 rounded-xl">
                  {lineasBorrador.map((l) => (
                    <div key={l.idLinea} className="px-3.5 py-2.5 flex items-center justify-between gap-3 text-xs">
                      <div className="min-w-0">
                        <div className="font-medium text-[#111111] truncate">{l.nombreProducto}</div>
                        <div className="text-[11px] text-zinc-400 tabular-nums">
                          {l.requiereRevision
                            ? `Pendiente definir (${l.cantidadTotal} u.)`
                            : l.desglose
                                .filter((d) => d.cantidad > 0)
                                .map((d) => `${d.codigo} ${d.etiqueta}: ${d.cantidad}`)
                                .join(' · ')}
                        </div>
                      </div>
                      <div className="flex items-center gap-3 shrink-0 tabular-nums">
                        <span className="font-semibold text-[#111111]">
                          {formatARS(round2(l.precioUnitarioNeto * l.cantidadTotal))}
                        </span>
                        <button
                          type="button"
                          onClick={() => handleEliminarLineaBorrador(l.idGrupo)}
                          className="p-1 text-zinc-400 hover:text-red-600"
                          aria-label="Quitar artículo"
                        >
                          <X className="w-3.5 h-3.5" strokeWidth={1.5} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>

                <div className="pt-2 space-y-1.5 text-xs tabular-nums">
                  <div className="flex justify-between text-zinc-500">
                    <span>Subtotal Neto</span>
                    <span>{formatARS(totalesBorrador.subtotalNeto)}</span>
                  </div>
                  {totalesBorrador.montoDescuento > 0 && (
                    <div className="flex justify-between text-zinc-500">
                      <span>Descuento ({form.descuentoPorc}%)</span>
                      <span>-{formatARS(totalesBorrador.montoDescuento)}</span>
                    </div>
                  )}
                  <div className="flex justify-between text-zinc-500">
                    <span>Base Imponible (Total B)</span>
                    <span>{formatARS(totalesBorrador.baseImponible)}</span>
                  </div>
                  <div className="flex justify-between text-zinc-500">
                    <span>IVA 21%</span>
                    <span>{formatARS(totalesBorrador.iva21)}</span>
                  </div>
                  <div className="flex justify-between text-sm font-semibold text-[#111111] pt-2 border-t border-zinc-200/70">
                    <span>Total Facturado</span>
                    <span>{formatARS(totalesBorrador.totalFacturado)}</span>
                  </div>
                </div>

                <button
                  type="button"
                  disabled={guardandoPedido || advertenciasExcel.length > 0}
                  onClick={() => void handleGuardarNuevoPedido()}
                  className="w-full py-3 rounded-full text-sm font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors disabled:opacity-50"
                >
                  {guardandoPedido ? 'Guardando...' : 'Guardar pedido'}
                </button>
              </section>
            )}
          </div>
        )}
      </main>

      <footer className="py-6 text-center text-[11px] text-zinc-400 border-t border-zinc-200/50">
        {MIBI_BRAND.pieOficial}
      </footer>

      {/* ==================================================================== */}
      {/* BARRA FLOTANTE CONTEXTUAL DE ACCIONES EN LOTE (APPLE / LINEAR STYLE) */}
      {/* ==================================================================== */}
      {vista === 'pedidos' && selectedIds.size > 0 && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 z-40 max-w-[96vw]">
          <div className="relative">
            {batchPanel === 'estado' && (
              <div className="absolute bottom-full mb-2.5 left-1/2 -translate-x-1/2 w-80 bg-white text-[#111111] rounded-2xl border border-[#EAEAEA] shadow-[0_12px_36px_rgba(0,0,0,0.12)] p-4 space-y-3.5">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-[#111111]">Armado y envío en lote</span>
                  <button
                    type="button"
                    onClick={() => setBatchPanel(null)}
                    className="p-1 rounded-full text-zinc-400 hover:text-zinc-800"
                  >
                    <X className="w-3.5 h-3.5" strokeWidth={1.5} />
                  </button>
                </div>

                <div className="space-y-1.5">
                  <span className="text-[11px] text-zinc-400 font-medium">Cambiar armado</span>
                  <div className="grid grid-cols-2 gap-1.5">
                    <button
                      type="button"
                      onClick={() => handleBatchArmado('En armado')}
                      className="px-3 py-2 rounded-xl text-xs font-medium bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 text-zinc-800 transition-colors"
                    >
                      En armado
                    </button>
                    <button
                      type="button"
                      onClick={() => handleBatchArmado('Listo')}
                      className="px-3 py-2 rounded-xl text-xs font-medium bg-emerald-50 hover:bg-emerald-100/80 border border-emerald-200/70 text-emerald-800 transition-colors"
                    >
                      Listo para armar
                    </button>
                  </div>
                </div>

                <div className="pt-3 border-t border-zinc-200/70 space-y-2">
                  <span className="text-[11px] text-zinc-400 font-medium">Despacho y transporte</span>
                  <input
                    type="text"
                    value={batchTransporte}
                    onChange={(e) => setBatchTransporte(e.target.value)}
                    placeholder="Transporte opcional (ej. Andreani, Cruz del Sur)"
                    className="w-full px-3 py-2 text-xs bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900"
                  />
                  <div className="grid grid-cols-2 gap-1.5">
                    <button
                      type="button"
                      onClick={() => handleBatchEnvio('Despachado')}
                      className="px-3 py-2 rounded-xl text-xs font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors"
                    >
                      Marcar Despachado
                    </button>
                    <button
                      type="button"
                      onClick={() => handleBatchEnvio('Entregado')}
                      className="px-3 py-2 rounded-xl text-xs font-medium bg-white hover:bg-zinc-50 border border-zinc-200 text-zinc-700 transition-colors"
                    >
                      Marcar Entregado
                    </button>
                  </div>
                </div>
              </div>
            )}

            {batchPanel === 'pago' && esAdmin && (
              <div className="absolute bottom-full mb-2.5 left-1/2 -translate-x-1/2 w-80 bg-white text-[#111111] rounded-2xl border border-[#EAEAEA] shadow-[0_12px_36px_rgba(0,0,0,0.12)] p-4 space-y-3.5">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-[#111111]">Pago y vendedor en lote</span>
                  <button
                    type="button"
                    onClick={() => setBatchPanel(null)}
                    className="p-1 rounded-full text-zinc-400 hover:text-zinc-800"
                  >
                    <X className="w-3.5 h-3.5" strokeWidth={1.5} />
                  </button>
                </div>

                <div className="space-y-1.5">
                  <span className="text-[11px] text-zinc-400 font-medium">Estado financiero (Administración)</span>
                  <button
                    type="button"
                    onClick={() => handleBatchPago('Pagado')}
                    className="w-full px-3 py-2 rounded-xl text-xs font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors"
                  >
                    Marcar Pagado (Cobrado en mano)
                  </button>
                  <button
                    type="button"
                    onClick={() => handleBatchPago('Falta pago')}
                    className="w-full px-3 py-2 rounded-xl text-xs font-medium bg-white hover:bg-red-50 border border-zinc-200 text-zinc-700 hover:text-red-700 transition-colors"
                  >
                    Pasar a Falta pago
                  </button>
                </div>

                <div className="pt-3 border-t border-zinc-200/70 space-y-2">
                  <span className="text-[11px] text-zinc-400 font-medium">Asignar vendedor al lote</span>
                  <div className="flex items-center gap-1.5">
                    <select
                      value={batchVendedor}
                      onChange={(e) => setBatchVendedor(e.target.value)}
                      className="flex-1 px-3 py-2 text-xs bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900"
                    >
                      {VENDEDORES_MIBI.map((v) => (
                        <option key={v} value={v}>
                          {v}
                        </option>
                      ))}
                    </select>
                    <button
                      type="button"
                      onClick={handleBatchVendedor}
                      className="px-3.5 py-2 rounded-xl text-xs font-medium bg-zinc-900 hover:bg-zinc-800 text-white"
                    >
                      Aplicar
                    </button>
                  </div>
                </div>
              </div>
            )}

            {batchPanel === 'exportar' && (
              <div className="absolute bottom-full mb-2.5 left-1/2 -translate-x-1/2 w-80 bg-white text-[#111111] rounded-2xl border border-[#EAEAEA] shadow-[0_12px_36px_rgba(0,0,0,0.12)] p-4 space-y-2.5">
                <div className="flex items-center justify-between pb-1">
                  <span className="text-xs font-semibold text-[#111111]">
                    Exportar {pedidosSeleccionados.length}{' '}
                    {pedidosSeleccionados.length === 1 ? 'seleccionado' : 'seleccionados'}
                  </span>
                  <button
                    type="button"
                    onClick={() => setBatchPanel(null)}
                    className="p-1 rounded-full text-zinc-400 hover:text-zinc-800"
                  >
                    <X className="w-3.5 h-3.5" strokeWidth={1.5} />
                  </button>
                </div>

                <button
                  type="button"
                  onClick={() => {
                    if (!XLSX) {
                      notify('Preparando Excel, intentá en un segundo.');
                      return;
                    }
                    const resLote = calcularResumenFinanciero(pedidosSeleccionados);
                    void exportarTableroMaestro(
                      pedidosSeleccionados,
                      resLote,
                      XLSX,
                      `SELECCION_${pedidosSeleccionados.length}`
                    ).then(() => {
                      notify(`Exportado Excel General (${pedidosSeleccionados.length} pedidos).`, 'green');
                    });
                    setBatchPanel(null);
                  }}
                  className="w-full text-left px-3.5 py-2.5 rounded-xl border border-zinc-200 hover:bg-zinc-50 transition-colors"
                >
                  <div className="text-xs font-semibold text-[#111111] flex items-center gap-1.5">
                    <FileSpreadsheet className="w-3.5 h-3.5 text-zinc-500" strokeWidth={1.5} />
                    Excel General (Seleccionados)
                  </div>
                  <p className="text-[11px] text-zinc-500 mt-0.5">
                    Planilla oficial de 17 columnas solo con los tildados.
                  </p>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    const res = exportarPickingConsolidado(pedidosSeleccionados, XLSX);
                    setBatchPanel(null);
                    if (!res.success) {
                      notify(res.error ?? 'No se pudo generar el consolidado.', 'red');
                      return;
                    }
                    const notas: string[] = [];
                    if (res.excluidosTalle.length > 0) {
                      notas.push(`excluidos sin talle: ${res.excluidosTalle.join(', ')}`);
                    }
                    if (res.excluidosSinLineas.length > 0) {
                      notas.push(`sin códigos: ${res.excluidosSinLineas.join(', ')}`);
                    }
                    notify(
                      notas.length > 0
                        ? `Consolidado listo (${res.incluidos} pedidos · ${res.totalUnidades} u. · ${notas.join(' · ')})`
                        : `Consolidado para Depósito listo (${res.incluidos} pedidos · ${res.totalUnidades} u.).`,
                      notas.length > 0 ? 'amber' : 'green'
                    );
                  }}
                  className="w-full text-left px-3.5 py-2.5 rounded-xl border border-zinc-900 bg-zinc-900 text-white hover:bg-zinc-800 transition-colors"
                >
                  <div className="text-xs font-semibold flex items-center gap-1.5">
                    <Layers className="w-3.5 h-3.5 text-[#F4B6CE]" strokeWidth={1.5} />
                    Consolidado para Depósito (Picking)
                  </div>
                  <p className="text-[11px] text-zinc-300 mt-0.5">
                    Suma unidades totales por código + pestaña por cliente.
                  </p>
                </button>
              </div>
            )}

            {batchPanel === 'eliminar' && esAdmin && (
              <div className="absolute bottom-full mb-2.5 right-0 w-72 bg-white text-[#111111] rounded-2xl border border-red-200/80 shadow-[0_12px_36px_rgba(0,0,0,0.12)] p-4 space-y-3">
                <div className="text-xs font-semibold text-red-700">
                  ¿Eliminar {pedidosSeleccionados.length}{' '}
                  {pedidosSeleccionados.length === 1 ? 'pedido' : 'pedidos'}?
                </div>
                <p className="text-[11px] text-zinc-500 tabular-nums">
                  Total involucrado: <strong className="text-[#111111]">{formatARS(totalSeleccionado)}</strong>. Tendrás
                  6 segundos para deshacer si fue por error.
                </p>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setBatchPanel(null)}
                    className="flex-1 py-2 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50"
                  >
                    Cancelar
                  </button>
                  <button
                    type="button"
                    onClick={handleConfirmarEliminarLote}
                    className="flex-1 py-2 rounded-full text-xs font-medium bg-red-600 hover:bg-red-700 text-white"
                  >
                    Confirmar
                  </button>
                </div>
              </div>
            )}

            {/* CÁPSULA FLOTANTE COMPACTA */}
            <div className="bg-zinc-900 text-white rounded-full shadow-lg px-3.5 py-2 flex items-center gap-2 sm:gap-3 border border-zinc-800">
              <div className="flex items-center gap-2 pl-1 whitespace-nowrap">
                <span className="inline-flex items-center justify-center min-w-[20px] h-5 px-1.5 rounded-full bg-[#F4B6CE] text-[#111111] text-[11px] font-bold tabular-nums">
                  {selectedIds.size}
                </span>
                <span className="text-xs font-medium hidden sm:inline">
                  {selectedIds.size === 1 ? 'seleccionado' : 'seleccionados'}
                </span>
                <span className="text-xs text-zinc-400 tabular-nums border-l border-zinc-700 pl-2">
                  {formatARS(totalSeleccionado)}
                </span>
              </div>

              <div className="h-4 w-px bg-zinc-700/80" />

              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => setBatchPanel((prev) => (prev === 'estado' ? null : 'estado'))}
                  className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-full text-xs font-medium transition-colors ${
                    batchPanel === 'estado' ? 'bg-white text-[#111111]' : 'text-zinc-200 hover:bg-zinc-800'
                  }`}
                >
                  <Truck className="w-3.5 h-3.5" strokeWidth={1.5} />
                  <span>Estado</span>
                </button>

                {esAdmin ? (
                  <button
                    type="button"
                    onClick={() => setBatchPanel((prev) => (prev === 'pago' ? null : 'pago'))}
                    className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-full text-xs font-medium transition-colors ${
                      batchPanel === 'pago' ? 'bg-white text-[#111111]' : 'text-zinc-200 hover:bg-zinc-800'
                    }`}
                  >
                    <CreditCard className="w-3.5 h-3.5" strokeWidth={1.5} />
                    <span>Pago</span>
                  </button>
                ) : (
                  <span
                    className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-full text-xs text-zinc-500 cursor-not-allowed"
                    title="Edición de pagos en lote exclusiva para Administración"
                  >
                    <Lock className="w-3 h-3" strokeWidth={1.5} />
                    <span className="hidden md:inline">Pago</span>
                  </span>
                )}

                <button
                  type="button"
                  onClick={() => setBatchPanel((prev) => (prev === 'exportar' ? null : 'exportar'))}
                  className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-full text-xs font-medium transition-colors ${
                    batchPanel === 'exportar' ? 'bg-white text-[#111111]' : 'text-zinc-200 hover:bg-zinc-800'
                  }`}
                >
                  <Download className="w-3.5 h-3.5" strokeWidth={1.5} />
                  <span>Exportar</span>
                </button>

                {esAdmin && (
                  <button
                    type="button"
                    onClick={() => setBatchPanel((prev) => (prev === 'eliminar' ? null : 'eliminar'))}
                    className={`inline-flex items-center gap-1 px-2.5 py-1.5 rounded-full text-xs font-medium transition-colors ${
                      batchPanel === 'eliminar'
                        ? 'bg-red-600 text-white'
                        : 'text-red-300 hover:bg-red-950/60 hover:text-red-200'
                    }`}
                  >
                    <Trash2 className="w-3.5 h-3.5" strokeWidth={1.5} />
                    <span className="hidden sm:inline">Eliminar ({selectedIds.size})</span>
                  </button>
                )}
              </div>

              <button
                type="button"
                onClick={limpiarSeleccion}
                className="p-1.5 rounded-full text-zinc-400 hover:text-white hover:bg-zinc-800 transition-colors"
                aria-label="Limpiar selección (Escape)"
                title="Deseleccionar (Esc)"
              >
                <X className="w-3.5 h-3.5" strokeWidth={1.75} />
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* PANEL LATERAL DERECHO (DRAWER DE PEDIDO)                             */}
      {/* ==================================================================== */}
      {pedidoDrawer && (
        <div
          className="fixed inset-0 z-50 flex justify-end bg-black/25 backdrop-blur-[1px]"
          onClick={() => {
            setDrawerPedidoId(null);
            setMostrarHojaImprimible(false);
          }}
        >
          <aside
            onClick={(e) => e.stopPropagation()}
            className="w-full max-w-xl bg-white h-full overflow-y-auto border-l border-[#EAEAEA] shadow-2xl flex flex-col"
          >
            {/* Cabecera del Drawer */}
            <div className="sticky top-0 z-10 bg-white/95 backdrop-blur border-b border-[#EAEAEA] px-6 py-4 flex items-center justify-between gap-3">
              <div>
                <div className="flex items-center gap-2">
                  <h2 className="text-base font-semibold text-[#111111] tabular-nums">
                    Pedido #{pedidoDrawer.numeroPedido}
                  </h2>
                  {pedidoDrawer.faltantesVariante > 0 ? (
                    <StatusPill tone="red">Falta talle o color</StatusPill>
                  ) : (
                    <StatusPill tone={pedidoDrawer.estadoArmado === 'Listo' ? 'green' : 'neutral'}>
                      {pedidoDrawer.estadoArmado}
                    </StatusPill>
                  )}
                </div>
                <p className="text-xs text-zinc-500 mt-0.5">
                  {pedidoDrawer.cliente} · {pedidoDrawer.vendedor}
                </p>
              </div>
              <button
                type="button"
                onClick={() => {
                  setDrawerPedidoId(null);
                  setMostrarHojaImprimible(false);
                }}
                className="p-2 rounded-full text-zinc-400 hover:text-zinc-900 hover:bg-zinc-100"
                aria-label="Cerrar panel"
              >
                <X className="w-4 h-4" strokeWidth={1.5} />
              </button>
            </div>

            <div className="p-6 space-y-6 flex-1">
              {/* Acciones rápidas */}
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  disabled={pedidoDrawer.faltantesVariante > 0 || pedidoDrawer.lineas.length === 0}
                  onClick={() => {
                    const res = exportarHojaDeposito(pedidoDrawer, XLSX);
                    if (!res.success) {
                      notify(res.errors[0] ?? 'No se pudo exportar.', 'red');
                      return;
                    }
                    notify(`Hoja depósito #${pedidoDrawer.numeroPedido} descargada.`, 'green');
                  }}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  <Download className="w-3.5 h-3.5" strokeWidth={1.5} />
                  Hoja depósito (.xlsx)
                </button>

                <button
                  type="button"
                  disabled={pedidoDrawer.faltantesVariante > 0}
                  onClick={() => setMostrarHojaImprimible(!mostrarHojaImprimible)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  <Printer className="w-3.5 h-3.5" strokeWidth={1.5} />
                  {mostrarHojaImprimible ? 'Ver gestión' : 'Vista imprimible'}
                </button>

                <button
                  type="button"
                  onClick={() => handleDuplicarComoNuevoPedido(pedidoDrawer)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors"
                  title="Crea un pedido nuevo prellenando cliente, vendedor y WhatsApp"
                >
                  <Copy className="w-3.5 h-3.5" strokeWidth={1.5} />
                  Duplicar como nuevo
                </button>
              </div>

              {mostrarHojaImprimible ? (
                <div className="mibi-print-area space-y-4 border border-zinc-200 rounded-2xl p-5">
                  <div className="flex items-start justify-between border-b border-zinc-200 pb-3">
                    <div>
                      <MibiLogo width={76} variant="pink" />
                      <div className="text-xs font-semibold mt-2">
                        HOJA PARA DEPÓSITO — PEDIDO #{pedidoDrawer.numeroPedido}
                      </div>
                      <div className="text-[11px] text-zinc-500">
                        Cliente: {pedidoDrawer.cliente} · Transporte: {pedidoDrawer.transporte || 'A definir'}
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => window.print()}
                      className="px-3 py-1.5 rounded-full text-xs font-medium bg-zinc-900 text-white"
                    >
                      Imprimir ahora
                    </button>
                  </div>

                  <table className="w-full text-xs border-collapse">
                    <thead>
                      <tr className="border-b border-zinc-200 text-zinc-500">
                        <th className="py-2 text-left">Código</th>
                        <th className="py-2 text-left">Artículo</th>
                        <th className="py-2 text-left">Talle / Color</th>
                        <th className="py-2 text-right">Cant.</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-200/60 tabular-nums">
                      {pedidoDrawer.lineas.flatMap((l) =>
                        l.desglose
                          .filter((d) => d.cantidad > 0)
                          .map((d) => (
                            <tr key={`${l.idLinea}-${d.codigo}`}>
                              <td className="py-2 font-semibold">{d.codigo}</td>
                              <td className="py-2">{l.nombreProducto}</td>
                              <td className="py-2">{d.etiqueta}</td>
                              <td className="py-2 text-right font-semibold">{d.cantidad}</td>
                            </tr>
                          ))
                      )}
                    </tbody>
                  </table>
                  <p className="text-[10px] text-zinc-400 pt-2 border-t border-zinc-200">{MIBI_BRAND.pieOficial}</p>
                </div>
              ) : (
                <>
                  {/* RESOLUCIÓN DE TALLES / COLORES PENDIENTES */}
                  {pedidoDrawer.faltantesVariante > 0 && (
                    <div className="bg-red-50/70 border border-red-200 rounded-2xl p-4 space-y-3">
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-xs font-semibold text-red-800">
                          Falta definir talle o color ({pedidoDrawer.faltantesVariante})
                        </span>
                        <button
                          type="button"
                          onClick={() => {
                            const msg = generarMensajeWhatsApp(
                              pedidoDrawer.cliente,
                              pedidoDrawer.numeroPedido,
                              pedidoDrawer.lineas
                            );
                            copyTextSafe(msg);
                            setCopiadoId(`wsp-${pedidoDrawer.id}`);
                            notify('Reclamo para WhatsApp copiado.', 'green');
                            window.setTimeout(() => setCopiadoId(null), 2200);
                          }}
                          className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-medium bg-white border border-red-200 text-red-700 hover:bg-red-50"
                        >
                          {copiadoId === `wsp-${pedidoDrawer.id}` ? (
                            <Check className="w-3 h-3 text-emerald-600" strokeWidth={1.75} />
                          ) : (
                            <MessageCircle className="w-3 h-3" strokeWidth={1.5} />
                          )}
                          {copiadoId === `wsp-${pedidoDrawer.id}` ? 'Copiado' : 'Copiar WhatsApp'}
                        </button>
                      </div>

                      {pedidoDrawer.lineas
                        .filter((l) => l.requiereRevision || l.estado !== 'OK')
                        .map((l) => {
                          const prod = CATALOGO_BY_GRUPO.get(l.idGrupo);
                          if (!prod || prod.variantes.length < 2) return null;
                          const [v0, v1] = prod.variantes;
                          const key = `drw_${pedidoDrawer.id}_${l.idLinea}`;
                          const cur = resolucionesManual[key] ?? { q0: '0', q1: '0' };
                          return (
                            <div key={l.idLinea} className="bg-white rounded-xl border border-red-200/70 p-3 space-y-2">
                              <div className="flex justify-between text-xs">
                                <span className="font-semibold">{l.nombreProducto}</span>
                                <span className="font-semibold text-red-700 tabular-nums">Total: {l.cantidadTotal} u.</span>
                              </div>
                              <div className="flex flex-wrap items-end gap-2">
                                <div className="flex-1 min-w-[100px]">
                                  <label className="block text-[11px] text-zinc-500 mb-1">
                                    {v0.etiqueta} ({v0.codigo})
                                  </label>
                                  <input
                                    type="number"
                                    min={0}
                                    max={l.cantidadTotal}
                                    value={cur.q0}
                                    onChange={(e) =>
                                      setResolucionesManual((prev) => ({
                                        ...prev,
                                        [key]: { ...cur, q0: e.target.value }
                                      }))
                                    }
                                    className="w-full px-2.5 py-1.5 text-xs bg-white border border-zinc-200 rounded-lg tabular-nums"
                                  />
                                </div>
                                <div className="flex-1 min-w-[100px]">
                                  <label className="block text-[11px] text-zinc-500 mb-1">
                                    {v1.etiqueta} ({v1.codigo})
                                  </label>
                                  <input
                                    type="number"
                                    min={0}
                                    max={l.cantidadTotal}
                                    value={cur.q1}
                                    onChange={(e) =>
                                      setResolucionesManual((prev) => ({
                                        ...prev,
                                        [key]: { ...cur, q1: e.target.value }
                                      }))
                                    }
                                    className="w-full px-2.5 py-1.5 text-xs bg-white border border-zinc-200 rounded-lg tabular-nums"
                                  />
                                </div>
                                <button
                                  type="button"
                                  onClick={() =>
                                    handleResolverEnDrawer(
                                      pedidoDrawer,
                                      l.idLinea,
                                      parseInt(cur.q0, 10) || 0,
                                      parseInt(cur.q1, 10) || 0
                                    )
                                  }
                                  className="px-3.5 py-1.5 rounded-full text-xs font-medium bg-zinc-900 text-white"
                                >
                                  Confirmar
                                </button>
                              </div>
                            </div>
                          );
                        })}
                    </div>
                  )}

                  {/* SECCIÓN PAGOS Y CHEQUES */}
                  <div className="space-y-3">
                    <div className="flex items-center justify-between">
                      <h3 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">
                        Pagos y liquidación
                      </h3>
                      {pedidoDrawer.estadoPago === 'Pagado' ? (
                        <StatusPill tone="green">Pagado</StatusPill>
                      ) : pedidoDrawer.estadoPago === 'Pendiente' ? (
                        <StatusPill tone="amber">Pendiente</StatusPill>
                      ) : (
                        <StatusPill tone="red">Falta pago</StatusPill>
                      )}
                    </div>

                    <div className="bg-[#FAFAFA] rounded-2xl border border-zinc-200/80 p-4 space-y-1.5 text-xs tabular-nums">
                      <div className="flex justify-between text-zinc-500">
                        <span>Subtotal Neto</span>
                        <span>{formatARS(pedidoDrawer.subtotalNeto)}</span>
                      </div>
                      {pedidoDrawer.descuentoPorc > 0 && (
                        <div className="flex justify-between text-zinc-500">
                          <span>Descuento ({pedidoDrawer.descuentoPorc}%)</span>
                          <span>-{formatARS(pedidoDrawer.montoDescuento)}</span>
                        </div>
                      )}
                      <div className="flex justify-between text-zinc-500">
                        <span>Base Imponible (Total B)</span>
                        <span>{formatARS(pedidoDrawer.baseImponible)}</span>
                      </div>
                      <div className="flex justify-between text-zinc-500">
                        <span>IVA 21%</span>
                        <span>{formatARS(pedidoDrawer.iva21)}</span>
                      </div>
                      <div className="flex justify-between font-semibold text-sm text-[#111111] pt-1.5 border-t border-zinc-200/70">
                        <span>Total Facturado</span>
                        <span>{formatARS(pedidoDrawer.totalFacturado)}</span>
                      </div>
                      <div className="flex justify-between text-emerald-700 font-medium">
                        <span>Cobrado en mano</span>
                        <span>{formatARS(pedidoDrawer.montoAbonado)}</span>
                      </div>
                      <div className="flex justify-between text-zinc-700 font-medium">
                        <span>Saldo pendiente</span>
                        <span>{formatARS(pedidoDrawer.saldoPendiente)}</span>
                      </div>
                    </div>

                    {esAdmin ? (
                      <div className="space-y-2.5 text-xs">
                        {pedidoDrawer.medioPago === 'Cheque' && !pedidoDrawer.chequeAcreditado && (
                          <button
                            type="button"
                            onClick={() => confirmarAcreditacion(pedidoDrawer)}
                            className="w-full py-2.5 rounded-full text-xs font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors"
                          >
                            Confirmar plata acreditada → Pasar a Pagado
                          </button>
                        )}

                        <div className="grid grid-cols-3 gap-2">
                          {(['Sin pagar', 'Transferencia / Efectivo', 'Cheque'] as const).map((m) => (
                            <button
                              key={m}
                              type="button"
                              onClick={() => {
                                if (m === 'Transferencia / Efectivo') {
                                  actualizarPedido(
                                    pedidoDrawer.id,
                                    {
                                      medioPago: m,
                                      montoAbonado: pedidoDrawer.totalFacturado,
                                      montoCheque: 0,
                                      chequeAcreditado: true
                                    },
                                    'Marcó pago completo (Transferencia / Efectivo)'
                                  );
                                } else if (m === 'Cheque') {
                                  const hoy = getTodayISO();
                                  actualizarPedido(
                                    pedidoDrawer.id,
                                    {
                                      medioPago: 'Cheque',
                                      chequeAcreditado: false,
                                      montoCheque: round2(pedidoDrawer.totalFacturado - pedidoDrawer.montoAbonado),
                                      fechaRecepcionCheque: pedidoDrawer.fechaRecepcionCheque || hoy,
                                      fechaAcreditacionCheque:
                                        pedidoDrawer.fechaAcreditacionCheque || addDaysISO(hoy, 30)
                                    },
                                    'Registró pago con Cheque (30 días)'
                                  );
                                } else {
                                  actualizarPedido(
                                    pedidoDrawer.id,
                                    {
                                      medioPago: 'Sin pagar',
                                      montoAbonado: 0,
                                      montoCheque: 0,
                                      chequeAcreditado: false
                                    },
                                    'Pasó estado de pago a Sin pagar'
                                  );
                                }
                              }}
                              className={`py-2 px-2 rounded-xl border text-[11px] font-medium transition-colors ${
                                pedidoDrawer.medioPago === m
                                  ? 'bg-zinc-900 text-white border-zinc-900'
                                  : 'bg-white text-zinc-700 border-zinc-200 hover:bg-zinc-50'
                              }`}
                            >
                              {m === 'Transferencia / Efectivo' ? 'Cobrado' : m}
                            </button>
                          ))}
                        </div>

                        <div className="grid grid-cols-2 gap-3">
                          <div>
                            <label className="block text-[11px] text-zinc-500 mb-1">Monto abonado ($)</label>
                            <input
                              type="number"
                              min={0}
                              max={pedidoDrawer.totalFacturado}
                              step="0.01"
                              value={drawerDraft?.montoAbonado ?? String(pedidoDrawer.montoAbonado)}
                              onChange={(e) =>
                                setDrawerDraft((prev) => (prev ? { ...prev, montoAbonado: e.target.value } : prev))
                              }
                              onBlur={() => {
                                const val = Math.max(
                                  0,
                                  Math.min(pedidoDrawer.totalFacturado, parseFloat(drawerDraft?.montoAbonado ?? '0') || 0)
                                );
                                if (Math.abs(val - pedidoDrawer.montoAbonado) > 0.005) {
                                  actualizarPedido(
                                    pedidoDrawer.id,
                                    { montoAbonado: round2(val) },
                                    `Actualizó monto abonado a ${formatARS(val)}`
                                  );
                                }
                              }}
                              className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl tabular-nums"
                            />
                          </div>

                          <div>
                            <label className="block text-[11px] text-zinc-500 mb-1">Descuento (%)</label>
                            <input
                              type="number"
                              min={0}
                              max={100}
                              step="0.5"
                              value={drawerDraft?.descuentoPorc ?? String(pedidoDrawer.descuentoPorc)}
                              onChange={(e) =>
                                setDrawerDraft((prev) => (prev ? { ...prev, descuentoPorc: e.target.value } : prev))
                              }
                              onBlur={() => {
                                const val = Math.max(
                                  0,
                                  Math.min(100, parseFloat(drawerDraft?.descuentoPorc ?? '0') || 0)
                                );
                                if (val !== pedidoDrawer.descuentoPorc) {
                                  actualizarPedido(
                                    pedidoDrawer.id,
                                    { descuentoPorc: val },
                                    `Actualizó descuento a ${val}%`
                                  );
                                }
                              }}
                              className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl tabular-nums"
                            />
                          </div>
                        </div>

                        {pedidoDrawer.medioPago === 'Cheque' && !pedidoDrawer.chequeAcreditado && (
                          <div className="grid grid-cols-2 gap-3 pt-1">
                            <div>
                              <label className="block text-[11px] text-zinc-500 mb-1">Fecha acreditación cheque</label>
                              <input
                                type="date"
                                value={drawerDraft?.fechaAcreditacionCheque ?? pedidoDrawer.fechaAcreditacionCheque}
                                onChange={(e) =>
                                  setDrawerDraft((prev) =>
                                    prev ? { ...prev, fechaAcreditacionCheque: e.target.value } : prev
                                  )
                                }
                                onBlur={() => {
                                  const f = drawerDraft?.fechaAcreditacionCheque ?? '';
                                  if (f && f !== pedidoDrawer.fechaAcreditacionCheque) {
                                    actualizarPedido(
                                      pedidoDrawer.id,
                                      { fechaAcreditacionCheque: f },
                                      `Actualizó vencimiento de cheque a ${formatShortDate(f)}`
                                    );
                                  }
                                }}
                                className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl tabular-nums"
                              />
                            </div>
                            <div>
                              <label className="block text-[11px] text-zinc-500 mb-1">Detalle cheque</label>
                              <input
                                type="text"
                                value={drawerDraft?.detalleCheque ?? pedidoDrawer.detalleCheque}
                                onChange={(e) =>
                                  setDrawerDraft((prev) => (prev ? { ...prev, detalleCheque: e.target.value } : prev))
                                }
                                onBlur={() => {
                                  const d = (drawerDraft?.detalleCheque ?? '').trim();
                                  if (d !== pedidoDrawer.detalleCheque) {
                                    actualizarPedido(
                                      pedidoDrawer.id,
                                      { detalleCheque: d },
                                      'Actualizó detalle de cheque'
                                    );
                                  }
                                }}
                                className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                              />
                            </div>
                          </div>
                        )}
                      </div>
                    ) : (
                      <p className="text-[11px] text-zinc-400 flex items-center gap-1.5">
                        <Lock className="w-3 h-3" strokeWidth={1.5} />
                        Los pagos y cheques solo pueden modificarse con perfil Administración.
                      </p>
                    )}
                  </div>

                  {/* SECCIÓN ARMADO Y ENVÍO (guarda al salir del campo con updateDoc parcial) */}
                  <div className="pt-5 border-t border-zinc-200/70 space-y-3">
                    <div className="flex items-center justify-between gap-2">
                      <h3 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">Armado y envío</h3>
                      {(drawerDraft?.transporte ?? pedidoDrawer.transporte).trim() &&
                        (drawerDraft?.nroSeguimiento ?? pedidoDrawer.nroSeguimiento).trim() && (
                          <button
                            type="button"
                            onClick={() => {
                              const pedidoConDraft: Pedido = {
                                ...pedidoDrawer,
                                transporte: (drawerDraft?.transporte ?? pedidoDrawer.transporte).trim(),
                                nroSeguimiento: (drawerDraft?.nroSeguimiento ?? pedidoDrawer.nroSeguimiento).trim()
                              };
                              const msg = generarMensajeDespachoWhatsApp(pedidoConDraft);
                              if (!msg) return;
                              copyTextSafe(msg);
                              setCopiadoId(`envio-${pedidoDrawer.id}`);
                              notify('Aviso de envío copiado para WhatsApp', 'green');
                              window.setTimeout(() => setCopiadoId(null), 2200);
                            }}
                            className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-medium bg-zinc-100 hover:bg-zinc-200/80 text-zinc-700 transition-colors"
                          >
                            {copiadoId === `envio-${pedidoDrawer.id}` ? (
                              <Check className="w-3 h-3 text-emerald-600" strokeWidth={1.75} />
                            ) : (
                              <Send className="w-3 h-3" strokeWidth={1.5} />
                            )}
                            {copiadoId === `envio-${pedidoDrawer.id}` ? 'Aviso copiado' : 'Copiar aviso de envío'}
                          </button>
                        )}
                    </div>

                    <div className="grid grid-cols-2 gap-3 text-xs">
                      <div>
                        <label className="block text-[11px] text-zinc-500 mb-1">Estado de armado</label>
                        {pedidoDrawer.faltantesVariante > 0 ? (
                          <div className="px-3 py-2 rounded-xl bg-red-50 text-red-700 border border-red-200/70 font-medium">
                            Falta talle o color
                          </div>
                        ) : (
                          <select
                            value={pedidoDrawer.estadoArmado}
                            onChange={(e) => {
                              const nuevoEstado = e.target.value as EstadoArmado;
                              actualizarPedido(
                                pedidoDrawer.id,
                                { estadoArmado: nuevoEstado },
                                `Cambió armado a ${nuevoEstado}`
                              );
                            }}
                            className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                          >
                            <option value="Por armar">Por armar</option>
                            <option value="En armado">En armado</option>
                            <option value="Listo">Listo para armar</option>
                          </select>
                        )}
                      </div>

                      <div>
                        <label className="block text-[11px] text-zinc-500 mb-1">Estado de envío</label>
                        <select
                          value={pedidoDrawer.estadoEnvio}
                          onChange={(e) => {
                            const nuevoEnvio = e.target.value as EstadoEnvio;
                            actualizarPedido(
                              pedidoDrawer.id,
                              { estadoEnvio: nuevoEnvio },
                              `Cambió envío a ${nuevoEnvio}`
                            );
                          }}
                          className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                        >
                          <option value="Pendiente">Pendiente</option>
                          <option value="Despachado">Despachado</option>
                          <option value="Entregado">Entregado</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-[11px] text-zinc-500 mb-1">Transporte</label>
                        <input
                          type="text"
                          value={drawerDraft?.transporte ?? pedidoDrawer.transporte}
                          onChange={(e) =>
                            setDrawerDraft((prev) => (prev ? { ...prev, transporte: e.target.value } : prev))
                          }
                          onBlur={() => {
                            const val = (drawerDraft?.transporte ?? '').trim();
                            if (val !== pedidoDrawer.transporte) {
                              actualizarPedido(pedidoDrawer.id, { transporte: val }, `Transporte: ${val || 'vacío'}`);
                            }
                          }}
                          placeholder="Ej. Andreani, Cruz del Sur"
                          className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                        />
                      </div>

                      <div>
                        <label className="block text-[11px] text-zinc-500 mb-1">N.º de seguimiento / guía</label>
                        <input
                          type="text"
                          value={drawerDraft?.nroSeguimiento ?? pedidoDrawer.nroSeguimiento}
                          onChange={(e) =>
                            setDrawerDraft((prev) => (prev ? { ...prev, nroSeguimiento: e.target.value } : prev))
                          }
                          onBlur={() => {
                            const val = (drawerDraft?.nroSeguimiento ?? '').trim();
                            if (val !== pedidoDrawer.nroSeguimiento) {
                              actualizarPedido(
                                pedidoDrawer.id,
                                { nroSeguimiento: val },
                                `Guía de seguimiento: ${val || 'vacía'}`
                              );
                            }
                          }}
                          placeholder="Código de guía"
                          className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl tabular-nums"
                        />
                      </div>

                      <div className="col-span-2">
                        <label className="block text-[11px] text-zinc-500 mb-1">Observaciones</label>
                        <textarea
                          rows={2}
                          value={drawerDraft?.observaciones ?? pedidoDrawer.observaciones}
                          onChange={(e) =>
                            setDrawerDraft((prev) => (prev ? { ...prev, observaciones: e.target.value } : prev))
                          }
                          onBlur={() => {
                            const val = (drawerDraft?.observaciones ?? '').trim();
                            if (val !== pedidoDrawer.observaciones) {
                              actualizarPedido(pedidoDrawer.id, { observaciones: val }, 'Actualizó observaciones');
                            }
                          }}
                          className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                        />
                      </div>
                    </div>
                  </div>

                  {/* ARTÍCULOS DEL PEDIDO */}
                  {pedidoDrawer.lineas.length > 0 && (
                    <div className="pt-5 border-t border-zinc-200/70 space-y-2.5">
                      <h3 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">
                        Detalle de artículos ({pedidoDrawer.lineas.length})
                      </h3>
                      <div className="divide-y divide-zinc-200/60 border border-zinc-200/80 rounded-xl">
                        {pedidoDrawer.lineas.map((l) => (
                          <div key={l.idLinea} className="px-3.5 py-2.5 text-xs flex justify-between gap-2">
                            <div>
                              <div className="font-medium text-[#111111]">{l.nombreProducto}</div>
                              <div className="text-[11px] text-zinc-400 tabular-nums">
                                {l.requiereRevision
                                  ? `Sin definir (${l.cantidadTotal} u.)`
                                  : l.desglose
                                      .filter((d) => d.cantidad > 0)
                                      .map((d) => `${d.codigo} · ${d.etiqueta}: ${d.cantidad}`)
                                      .join(' | ')}
                              </div>
                            </div>
                            <div className="font-semibold tabular-nums whitespace-nowrap">
                              {formatARS(round2(l.precioUnitarioNeto * l.cantidadTotal))}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* HISTORIAL DE AUDITORÍA */}
                  <div className="pt-5 border-t border-zinc-200/70 space-y-2.5">
                    <h3 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">
                      Historial de cambios
                    </h3>
                    <ul className="space-y-1.5 text-[11px] text-zinc-500">
                      {pedidoDrawer.historial.map((h, i) => (
                        <li key={i} className="flex items-start gap-2">
                          <Clock className="w-3 h-3 text-zinc-400 mt-0.5 shrink-0" strokeWidth={1.5} />
                          <span>
                            <strong className="text-zinc-700 font-medium">{h.fechaHora}</strong> · {h.usuario} ({h.rol}):{' '}
                            {h.accion}
                          </span>
                        </li>
                      ))}
                    </ul>
                  </div>

                  {/* ELIMINACIÓN INDIVIDUAL (SOLO ADMINISTRACIÓN) */}
                  {esAdmin && (
                    <div className="pt-5 border-t border-zinc-200/70 flex justify-end">
                      {confirmEliminarId === pedidoDrawer.id ? (
                        <div className="flex items-center gap-2">
                          <button
                            type="button"
                            onClick={() => setConfirmEliminarId(null)}
                            className="px-3 py-1.5 rounded-full text-xs border border-zinc-200 text-zinc-600"
                          >
                            Cancelar
                          </button>
                          <button
                            type="button"
                            onClick={() => eliminarPedidoIndividual(pedidoDrawer)}
                            className="px-3.5 py-1.5 rounded-full text-xs font-medium bg-red-600 text-white"
                          >
                            Confirmar eliminación
                          </button>
                        </div>
                      ) : (
                        <button
                          type="button"
                          onClick={() => setConfirmEliminarId(pedidoDrawer.id)}
                          className="inline-flex items-center gap-1.5 text-xs text-red-600 hover:text-red-700"
                        >
                          <Trash2 className="w-3.5 h-3.5" strokeWidth={1.5} />
                          Eliminar pedido
                        </button>
                      )}
                    </div>
                  )}
                </>
              )}
            </div>
          </aside>
        </div>
      )}

      {/* ==================================================================== */}
      {/* MODAL DE EQUIPO, CLAVES Y VERIFICACIÓN QA EN VIVO (27+ PRUEBAS)      */}
      {/* ==================================================================== */}
      {modalEquipoAbierto && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/30 px-4"
          onClick={() => setModalEquipoAbierto(false)}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="w-full max-w-lg bg-white rounded-2xl border border-[#EAEAEA] shadow-2xl p-6 space-y-5 max-h-[88vh] overflow-y-auto"
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <ShieldCheck className="w-4 h-4 text-zinc-700" strokeWidth={1.5} />
                <h2 className="text-sm font-semibold text-[#111111]">Equipo y verificación técnica</h2>
              </div>
              <button
                type="button"
                onClick={() => setModalEquipoAbierto(false)}
                className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-900"
              >
                <X className="w-4 h-4" strokeWidth={1.5} />
              </button>
            </div>

            <div className="bg-[#FAFAFA] rounded-xl border border-zinc-200/80 px-4 py-3 text-xs flex items-center justify-between">
              <div>
                <div className="font-semibold text-[#111111]">{sesion?.nombre}</div>
                <div className="text-zinc-500">Perfil activo: {sesion?.rol}</div>
              </div>
              <StatusPill tone={esAdmin ? 'green' : 'neutral'}>{sesion?.rol ?? 'Equipo'}</StatusPill>
            </div>

            {esAdmin && (
              <form onSubmit={handleActualizarClavesEquipo} className="space-y-3 pt-2 border-t border-zinc-200/70">
                <div className="text-xs font-semibold text-zinc-600">Actualizar claves del equipo</div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-[11px] text-zinc-500 mb-1">Nueva clave Administración</label>
                    <input
                      type="password"
                      value={pinAdminModal}
                      onChange={(e) => setPinAdminModal(e.target.value)}
                      placeholder="Dejar vacío para no cambiar"
                      className="w-full px-3 py-2 text-xs bg-white border border-zinc-200 rounded-xl"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] text-zinc-500 mb-1">Nueva clave Ventas / Depósito</label>
                    <input
                      type="password"
                      value={pinVentasModal}
                      onChange={(e) => setPinVentasModal(e.target.value)}
                      placeholder="Dejar vacío para no cambiar"
                      className="w-full px-3 py-2 text-xs bg-white border border-zinc-200 rounded-xl"
                    />
                  </div>
                </div>
                <button
                  type="submit"
                  disabled={guardandoAcceso}
                  className="px-4 py-2 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800"
                >
                  Guardar nuevas claves
                </button>
              </form>
            )}

            {/* VERIFICACIÓN QA EN VIVO (27+ PRUEBAS SOBRE EL MOTOR REAL) */}
            <div className="pt-4 border-t border-zinc-200/70 space-y-2.5">
              <button
                type="button"
                onClick={() => setResultadosQA(ejecutarSuiteQA(XLSX))}
                className="w-full inline-flex items-center justify-center gap-1.5 py-2 rounded-full text-xs font-medium bg-[#FAFAFA] border border-zinc-200 text-zinc-700 hover:bg-zinc-100 transition-colors"
              >
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" strokeWidth={1.5} />
                Ejecutar verificación QA del motor (27+ pruebas)
              </button>

              {resultadosQA && (
                <div className="space-y-1.5 max-h-60 overflow-y-auto border border-zinc-200/80 rounded-xl p-3 bg-[#FAFAFA]">
                  <div className="text-xs font-semibold text-[#111111] pb-1 border-b border-zinc-200/70 tabular-nums">
                    Resultado: {resultadosQA.filter((r) => r.ok).length} / {resultadosQA.length} aprobadas
                  </div>
                  {resultadosQA.map((r, idx) => (
                    <div key={idx} className="flex items-start justify-between gap-2 text-[11px]">
                      <span className={r.ok ? 'text-zinc-700' : 'text-red-700 font-semibold'}>
                        {r.ok ? '✓' : '✗'} {r.nombre}
                      </span>
                      <span className="text-zinc-400 shrink-0 tabular-nums">{r.detalle}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
