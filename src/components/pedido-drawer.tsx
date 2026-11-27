'use client';

import React, { useEffect, useId, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  AlertTriangle,
  Check,
  Clock,
  Copy,
  Download,
  Edit3,
  Lock,
  MessageCircle,
  Minus,
  MoreHorizontal,
  Plus,
  Printer,
  Search,
  Send,
  Trash2,
  UserCheck,
  X
} from 'lucide-react';
import type { EnviarCuentasResult } from '@/lib/dal/pedidos';
import { normalizarTexto, resolverLinea } from '@/lib/excel-parser';
import {
  addDaysISO,
  formatARS,
  formatARSCompacto,
  formatShortDate,
  fromCents,
  getTodayISO,
  recalcularPedido,
  round2,
  toCents
} from '@/lib/finanzas';
import {
  isDepositoRole,
  type DrawerFocusTarget,
  type EstadoArmado,
  type EstadoEnvio,
  type LineaPedido,
  type Pedido,
  type PrioridadAsignacion,
  type ProductoCatalogo,
  type ResumenAjusteArticulos,
  type ToneStatus,
  type UsuarioEquipo
} from '@/lib/types';
import {
  buildWhatsAppUrl,
  generarMensajeAjusteStockWhatsApp,
  generarMensajeDespachoWhatsApp,
  generarMensajeWhatsApp
} from '@/lib/whatsapp';
import { MibiLogo } from './mibi-logo';
import { StatusPill } from './status-pill';

interface DrawerDraft {
  cliente: string;
  whatsapp: string;
  descuentoPorc: string;
  montoAbonado: string;
  fechaRecepcionCheque: string;
  fechaAcreditacionCheque: string;
  detalleCheque: string;
  transporte: string;
  nroSeguimiento: string;
  observaciones: string;
}

interface PedidoDrawerProps {
  pedido: Pedido;
  esAdmin: boolean;
  esDeposito: boolean;
  catalogoCompleto: ReadonlyArray<ProductoCatalogo>;
  cuentasAsignables: ReadonlyArray<UsuarioEquipo>;
  transportesSugeridos: ReadonlyArray<string>;
  focusTarget: DrawerFocusTarget;
  pendingKeys: ReadonlySet<string>;
  copiadoId: string | null;
  onClose: () => void;
  onPatch: (
    pedidoId: string,
    patch: Record<string, unknown>,
    motivo: string,
    expectedVersion?: number
  ) => Promise<boolean>;
  onConfirmarAcreditacion: (pedido: Pedido) => Promise<void>;
  onResolverLinea: (pedido: Pedido, idLinea: string, cantidades: number[]) => Promise<boolean>;
  onActualizarArticulos: (params: {
    pedidoId: string;
    expectedVersion: number;
    lineas: LineaPedido[];
    resolverIncidencia?: boolean;
    confirmarSobreDespachado?: boolean;
  }) => Promise<ResumenAjusteArticulos | null>;
  onExportarHojaDeposito: (pedido: Pedido) => void;
  onDuplicarComoNuevo: (pedido: Pedido) => void;
  onEliminarIndividual: (pedido: Pedido) => void;
  onEnviarACuentas: (params: {
    pedidoIds: string[];
    usuarioIds: string[];
    prioridad: PrioridadAsignacion;
    notaEnvio: string;
    forzarSinPago?: boolean;
  }) => Promise<EnviarCuentasResult>;
  onRetirarAsignacion: (pedidoId: string, usuarioId: string) => Promise<void>;
  onReportarIncidencia: (pedidoId: string, mensaje: string) => Promise<void>;
  onResolverIncidencia: (pedidoId: string) => Promise<void>;
  onCopyText: (key: string, text: string, mensajeOk: string) => void;
  onNotify: (texto: string, tone?: ToneStatus) => void;
}

const formatHoraCorta = (ts: number | null): string => {
  if (!ts) return '';
  const d = new Date(ts);
  return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
};

const QUICK_INCIDENCIA_CHIPS = ['Falta stock', 'Artículo dañado', 'Falta dato de envío'] as const;

type SaveSection = 'cliente' | 'pagos' | 'logistica';
type SaveStatus = 'saving' | 'saved';

export function PedidoDrawer({
  pedido,
  esAdmin,
  esDeposito,
  catalogoCompleto,
  cuentasAsignables,
  transportesSugeridos,
  focusTarget,
  pendingKeys,
  copiadoId,
  onClose,
  onPatch,
  onConfirmarAcreditacion,
  onResolverLinea,
  onActualizarArticulos,
  onExportarHojaDeposito,
  onDuplicarComoNuevo,
  onEliminarIndividual,
  onEnviarACuentas,
  onRetirarAsignacion,
  onReportarIncidencia,
  onResolverIncidencia,
  onCopyText,
  onNotify
}: PedidoDrawerProps): React.JSX.Element {
  const datalistId = useId();
  const [mostrarHojaImprimible, setMostrarHojaImprimible] = useState<boolean>(false);
  const [menuMasAbierto, setMenuMasAbierto] = useState<boolean>(false);
  const [confirmEliminar, setConfirmEliminar] = useState<boolean>(false);
  const [resolucionesManual, setResolucionesManual] = useState<Record<string, string[]>>({});
  const [editandoLineaId, setEditandoLineaId] = useState<string | null>(null);
  const [editandoPago, setEditandoPago] = useState<boolean>(false);
  const [cambiandoFormaCheque, setCambiandoFormaCheque] = useState<boolean>(false);
  const [confirmArmadoDestino, setConfirmArmadoDestino] = useState<EstadoArmado | null>(null);
  const [saveState, setSaveState] = useState<Partial<Record<SaveSection, SaveStatus>>>({});

  const [draft, setDraft] = useState<DrawerDraft>({
    cliente: pedido.cliente,
    whatsapp: pedido.whatsapp,
    descuentoPorc: String(pedido.descuentoPorc),
    montoAbonado: String(pedido.montoAbonado),
    fechaRecepcionCheque: pedido.fechaRecepcionCheque,
    fechaAcreditacionCheque: pedido.fechaAcreditacionCheque,
    detalleCheque: pedido.detalleCheque,
    transporte: pedido.transporte,
    nroSeguimiento: pedido.nroSeguimiento,
    observaciones: pedido.observaciones
  });

  // Derivación a cuentas (solo ADMIN)
  const [mostrarEnviar, setMostrarEnviar] = useState<boolean>(false);
  const [cuentaDestinoId, setCuentaDestinoId] = useState<string>('');
  const [prioridadEnvio, setPrioridadEnvio] = useState<PrioridadAsignacion>('NORMAL');
  const [notaEnvio, setNotaEnvio] = useState<string>('');
  const [requiereConfirmSinPago, setRequiereConfirmSinPago] = useState<boolean>(false);
  const [enviandoCuenta, setEnviandoCuenta] = useState<boolean>(false);

  // Reporte de incidencia (DEPOSITO)
  const [mostrarIncidenciaForm, setMostrarIncidenciaForm] = useState<boolean>(false);
  const [incidenciaTexto, setIncidenciaTexto] = useState<string>('');
  const [enviandoIncidencia, setEnviandoIncidencia] = useState<boolean>(false);

  // Modo "Ajustar cantidades" (reemplazo / faltantes de stock)
  const [editandoArticulos, setEditandoArticulos] = useState<boolean>(false);
  const [lineasEdicion, setLineasEdicion] = useState<LineaPedido[]>([]);
  const [mostrarBuscadorReemplazo, setMostrarBuscadorReemplazo] = useState<boolean>(false);
  const [busquedaReemplazo, setBusquedaReemplazo] = useState<string>('');
  const [reemplazoPendienteProd, setReemplazoPendienteProd] = useState<ProductoCatalogo | null>(null);
  const [reemplazoVarsQty, setReemplazoVarsQty] = useState<Record<string, string>>({});
  const [mostrarAdhocDrawer, setMostrarAdhocDrawer] = useState<boolean>(false);
  const [adhocNombre, setAdhocNombre] = useState<string>('');
  const [adhocRubro, setAdhocRubro] = useState<string>('PILUSOS');
  const [adhocPrecio, setAdhocPrecio] = useState<string>('');
  const [adhocCantidad, setAdhocCantidad] = useState<string>('1');
  const [confirmQuitarIdGrupo, setConfirmQuitarIdGrupo] = useState<string | null>(null);
  const [resolverIncidenciaCheck, setResolverIncidenciaCheck] = useState<boolean>(false);
  const [confirmSobreDespachadoCheck, setConfirmSobreDespachadoCheck] = useState<boolean>(false);
  const [ultimoAjusteGuardado, setUltimoAjusteGuardado] = useState<ResumenAjusteArticulos | null>(null);

  const tallesSectionRef = useRef<HTMLDivElement | null>(null);
  const firstTalleInputRef = useRef<HTMLInputElement | null>(null);
  const transporteInputRef = useRef<HTMLInputElement | null>(null);
  const derivacionSectionRef = useRef<HTMLDivElement | null>(null);
  const moreMenuRef = useRef<HTMLDivElement | null>(null);

  const cuentasDeposito = useMemo(
    () => cuentasAsignables.filter((c) => isDepositoRole(c.rol)),
    [cuentasAsignables]
  );
  const cuentasVentas = useMemo(
    () => cuentasAsignables.filter((c) => !isDepositoRole(c.rol)),
    [cuentasAsignables]
  );

  useEffect(() => {
    setEditandoArticulos(false);
    setLineasEdicion([]);
    setMostrarBuscadorReemplazo(false);
    setBusquedaReemplazo('');
    setReemplazoPendienteProd(null);
    setReemplazoVarsQty({});
    setMostrarAdhocDrawer(false);
    setConfirmQuitarIdGrupo(null);
    setConfirmSobreDespachadoCheck(false);
    setUltimoAjusteGuardado(null);
  }, [pedido.id]);

  useEffect(() => {
    setConfirmEliminar(false);
    setMostrarHojaImprimible(false);
    setMenuMasAbierto(false);
    setEditandoLineaId(null);
    setEditandoPago(false);
    setCambiandoFormaCheque(false);
    setConfirmArmadoDestino(null);
    setRequiereConfirmSinPago(false);
    setMostrarIncidenciaForm(false);
    setIncidenciaTexto('');
    setSaveState({});
    setMostrarEnviar(focusTarget === 'derivacion');
    if (cuentasDeposito.length === 1 && cuentasDeposito[0]) {
      setCuentaDestinoId(cuentasDeposito[0].id);
    } else {
      setCuentaDestinoId('');
    }
  }, [pedido.id, focusTarget, cuentasDeposito]);

  useEffect(() => {
    if (typeof document === 'undefined') return;
    if (mostrarHojaImprimible) {
      document.body.classList.add('mibi-printing-hoja');
    } else {
      document.body.classList.remove('mibi-printing-hoja');
    }
    return () => {
      document.body.classList.remove('mibi-printing-hoja');
    };
  }, [mostrarHojaImprimible]);

  useEffect(() => {
    if (!focusTarget) return;
    const t = window.setTimeout(() => {
      if (focusTarget === 'talles') {
        tallesSectionRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
        firstTalleInputRef.current?.focus();
        firstTalleInputRef.current?.select();
      } else if (focusTarget === 'transporte') {
        transporteInputRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
        transporteInputRef.current?.focus();
      } else if (focusTarget === 'derivacion') {
        derivacionSectionRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }, 60);
    return () => window.clearTimeout(t);
  }, [focusTarget, pedido.id]);

  useEffect(() => {
    if (!confirmEliminar) return;
    const t = window.setTimeout(() => setConfirmEliminar(false), 5000);
    return () => window.clearTimeout(t);
  }, [confirmEliminar]);

  useEffect(() => {
    if (!menuMasAbierto) return;
    const onDown = (e: MouseEvent): void => {
      if (moreMenuRef.current && !moreMenuRef.current.contains(e.target as Node)) {
        setMenuMasAbierto(false);
      }
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [menuMasAbierto]);

  useEffect(() => {
    setDraft({
      cliente: pedido.cliente,
      whatsapp: pedido.whatsapp,
      descuentoPorc: String(pedido.descuentoPorc),
      montoAbonado: String(pedido.montoAbonado),
      fechaRecepcionCheque: pedido.fechaRecepcionCheque,
      fechaAcreditacionCheque: pedido.fechaAcreditacionCheque,
      detalleCheque: pedido.detalleCheque,
      transporte: pedido.transporte,
      nroSeguimiento: pedido.nroSeguimiento,
      observaciones: pedido.observaciones
    });
  }, [
    pedido.id,
    pedido.version,
    pedido.cliente,
    pedido.whatsapp,
    pedido.descuentoPorc,
    pedido.montoAbonado,
    pedido.fechaRecepcionCheque,
    pedido.fechaAcreditacionCheque,
    pedido.detalleCheque,
    pedido.transporte,
    pedido.nroSeguimiento,
    pedido.observaciones
  ]);

  const patchConEstado = async (
    section: SaveSection,
    patch: Record<string, unknown>,
    motivo: string
  ): Promise<boolean> => {
    setSaveState((prev) => ({ ...prev, [section]: 'saving' }));
    const ok = await onPatch(pedido.id, patch, motivo, pedido.version);
    if (ok) {
      setSaveState((prev) => ({ ...prev, [section]: 'saved' }));
      window.setTimeout(() => {
        setSaveState((prev) => {
          if (prev[section] !== 'saved') return prev;
          const next = { ...prev };
          delete next[section];
          return next;
        });
      }, 2000);
    } else {
      setSaveState((prev) => {
        const next = { ...prev };
        delete next[section];
        return next;
      });
    }
    return ok;
  };

  const renderSaveBadge = (section: SaveSection): React.JSX.Element | null => {
    const st = saveState[section];
    if (!st) return null;
    if (st === 'saving') {
      return <span className="text-[11px] text-zinc-400 dark:text-zinc-500 font-normal">Guardando...</span>;
    }
    return (
      <span className="inline-flex items-center gap-1 text-[11px] text-emerald-700 dark:text-emerald-400 font-medium">
        <Check className="w-3 h-3" strokeWidth={2} /> Guardado
      </span>
    );
  };

  const commitCliente = (): void => {
    const val = draft.cliente.trim();
    if (!val) {
      setDraft((prev) => ({ ...prev, cliente: pedido.cliente }));
      onNotify('El nombre del cliente no puede quedar vacío.', 'amber');
      return;
    }
    if (val !== pedido.cliente) {
      const motivo = `Actualizó cliente: ${pedido.cliente} → ${val}`.slice(0, 300);
      void patchConEstado('cliente', { cliente: val }, motivo);
    }
  };

  const commitWhatsapp = (): void => {
    const val = draft.whatsapp.trim();
    if (val !== pedido.whatsapp) {
      void patchConEstado('cliente', { whatsapp: val }, `Actualizó WhatsApp: ${val || 'vacío'}`);
    }
  };

  const handleConfirmarEnvioCuenta = async (forzarSinPago = false): Promise<void> => {
    if (!cuentaDestinoId || enviandoCuenta) {
      if (!cuentaDestinoId) onNotify('Elegí una cuenta de destino.', 'amber');
      return;
    }
    setEnviandoCuenta(true);
    const res = await onEnviarACuentas({
      pedidoIds: [pedido.id],
      usuarioIds: [cuentaDestinoId],
      prioridad: prioridadEnvio,
      notaEnvio,
      forzarSinPago
    });
    setEnviandoCuenta(false);

    if (res.requiereConfirmacionSinPago) {
      setRequiereConfirmSinPago(true);
      return;
    }
    if (!res.ok) {
      onNotify(res.error ?? 'No se pudo enviar a la cuenta seleccionada.', 'red');
      return;
    }
    setRequiereConfirmSinPago(false);
    setMostrarEnviar(false);
    setNotaEnvio('');
    setPrioridadEnvio('NORMAL');
    onNotify(`Pedido #${pedido.numeroPedido} enviado a la cuenta.`, 'green');
  };

  const handleEnviarIncidencia = async (): Promise<void> => {
    if (enviandoIncidencia) return;
    const msg = incidenciaTexto.trim();
    if (!msg) {
      onNotify('Describí el faltante o problema antes de enviar.', 'amber');
      return;
    }
    setEnviandoIncidencia(true);
    await onReportarIncidencia(pedido.id, msg);
    setEnviandoIncidencia(false);
    setMostrarIncidenciaForm(false);
    setIncidenciaTexto('');
  };

  const handleCambioArmado = (destino: EstadoArmado): void => {
    if (destino === pedido.estadoArmado) return;
    if (pedido.faltantesVariante > 0 && destino === 'Listo') {
      onNotify('Definí talle o color antes de marcar como Listo.', 'red');
      return;
    }
    if (pedido.estadoEnvio !== 'Pendiente' && destino !== 'Listo' && confirmArmadoDestino !== destino) {
      setConfirmArmadoDestino(destino);
      return;
    }
    setConfirmArmadoDestino(null);
    void patchConEstado('logistica', { estadoArmado: destino }, `Cambió armado a ${destino}`);
  };

  const handleCambioEnvio = (destino: EstadoEnvio): void => {
    if (destino === pedido.estadoEnvio) return;
    if (pedido.faltantesVariante > 0 && destino !== 'Pendiente') {
      onNotify('Definí talle o color antes de despachar.', 'red');
      return;
    }
    void patchConEstado('logistica', { estadoEnvio: destino }, `Cambió envío a ${destino}`).then((ok) => {
      if (ok && pedido.estadoEnvio === 'Entregado' && destino !== 'Entregado') {
        onNotify(
          destino === 'Pendiente'
            ? 'Envío vuelto a Pendiente; se limpiaron las fechas de despacho y entrega.'
            : 'Envío vuelto a Despachado; se limpió la fecha de entrega.',
          'neutral'
        );
      }
    });
  };

  const renderVariantEditor = (
    l: LineaPedido,
    mode: 'pendiente' | 'edicion',
    isFirstCard: boolean
  ): React.JSX.Element | null => {
    const prod = catalogoCompleto.find((p) => p.idGrupo === l.idGrupo);
    const variantes = prod ? prod.variantes : l.desglose;
    if (variantes.length < 2) return null;

    const key = mode === 'pendiente' ? `drw_${pedido.id}_${l.idLinea}` : `ed_${pedido.id}_${l.idLinea}`;
    const defaultArr =
      mode === 'edicion'
        ? variantes.map((v) => String(l.desglose.find((d) => d.codigo === v.codigo)?.cantidad ?? 0))
        : variantes.map(() => '0');
    const cur = resolucionesManual[key] ?? defaultArr;
    const nums = variantes.map((_, idx) => Math.max(0, parseInt(cur[idx] ?? '0', 10) || 0));
    const suma = nums.reduce((acc, q) => acc + q, 0);
    const diff = l.cantidadTotal - suma;
    const esDos = variantes.length === 2;
    const busyKey = `talles:${pedido.id}:${l.idLinea}`;
    const busy = pendingKeys.has(busyKey);

    const setArr = (nextArr: string[]): void => {
      setResolucionesManual((prev) => ({ ...prev, [key]: nextArr }));
    };

    const confirmar = (): void => {
      if (suma !== l.cantidadTotal) {
        onNotify(`La suma debe ser exactamente ${l.cantidadTotal} u.`, 'red');
        return;
      }
      void onResolverLinea(pedido, l.idLinea, nums).then((ok) => {
        if (ok && mode === 'edicion') setEditandoLineaId(null);
      });
    };

    const handleVariantInput = (vIdx: number, raw: string): void => {
      if (esDos) {
        const cleaned = raw === '' ? 0 : Math.max(0, Math.min(l.cantidadTotal, parseInt(raw, 10) || 0));
        const other = Math.max(0, l.cantidadTotal - cleaned);
        if (vIdx === 0) {
          setArr([raw === '' ? '0' : String(cleaned), String(other)]);
        } else {
          setArr([String(other), raw === '' ? '0' : String(cleaned)]);
        }
        return;
      }
      setArr(variantes.map((_, idx) => (idx === vIdx ? raw : (cur[idx] ?? '0'))));
    };

    return (
      <div
        key={l.idLinea}
        className={`rounded-xl border p-3 space-y-2.5 ${
          mode === 'pendiente'
            ? 'bg-white dark:bg-[#1B1B1F] border-red-200/70 dark:border-red-500/25'
            : 'bg-[#FAFAFA] dark:bg-[#1B1B1F] border-zinc-200 dark:border-[#2E2E35]'
        }`}
      >
        <div className="flex justify-between items-center text-xs gap-2">
          <span className="font-semibold text-[#111111] dark:text-[#F4F4F5]">{l.nombreProducto}</span>
          <span
            className={`font-semibold tabular-nums ${
              mode === 'pendiente' ? 'text-red-700 dark:text-red-300' : 'text-zinc-600 dark:text-zinc-300'
            }`}
          >
            Total: {l.cantidadTotal} u.
          </span>
        </div>

        {mode === 'pendiente' && l.motivoRevision && (
          <p className="text-[11px] text-red-600 dark:text-red-400">{l.motivoRevision}</p>
        )}

        {mode === 'edicion' && pedido.estadoEnvio !== 'Pendiente' && (
          <p className="text-[11px] text-amber-700 dark:text-amber-400">
            Atención: este pedido ya está {pedido.estadoEnvio}. Si modificás talles, Depósito recibirá aviso de relectura.
          </p>
        )}

        {esDos && (
          <div className="flex flex-wrap items-center gap-1.5">
            <button
              type="button"
              onClick={() => setArr([String(l.cantidadTotal), '0'])}
              className="px-2.5 py-1 rounded-full text-[11px] font-medium bg-zinc-100 hover:bg-zinc-200/80 text-zinc-700 dark:bg-[#141417] dark:hover:bg-[#24242B] dark:text-zinc-300 border border-zinc-200/70 dark:border-[#2E2E35]"
            >
              Todo {variantes[0]?.etiqueta}
            </button>
            {l.cantidadTotal % 2 === 0 && (
              <button
                type="button"
                onClick={() => setArr([String(l.cantidadTotal / 2), String(l.cantidadTotal / 2)])}
                className="px-2.5 py-1 rounded-full text-[11px] font-medium bg-zinc-100 hover:bg-zinc-200/80 text-zinc-700 dark:bg-[#141417] dark:hover:bg-[#24242B] dark:text-zinc-300 border border-zinc-200/70 dark:border-[#2E2E35]"
              >
                Mitad y mitad
              </button>
            )}
            <button
              type="button"
              onClick={() => setArr(['0', String(l.cantidadTotal)])}
              className="px-2.5 py-1 rounded-full text-[11px] font-medium bg-zinc-100 hover:bg-zinc-200/80 text-zinc-700 dark:bg-[#141417] dark:hover:bg-[#24242B] dark:text-zinc-300 border border-zinc-200/70 dark:border-[#2E2E35]"
            >
              Todo {variantes[1]?.etiqueta}
            </button>
          </div>
        )}

        <div className="flex flex-wrap items-end gap-2">
          {variantes.map((v, vIdx) => (
            <div key={v.codigo} className="flex-1 min-w-[100px]">
              <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                {v.etiqueta} ({v.codigo})
              </label>
              <input
                ref={mode === 'pendiente' && isFirstCard && vIdx === 0 ? firstTalleInputRef : undefined}
                type="number"
                min={0}
                max={l.cantidadTotal}
                value={cur[vIdx] ?? '0'}
                onWheel={(e) => e.currentTarget.blur()}
                onChange={(e) => handleVariantInput(vIdx, e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    confirmar();
                  }
                }}
                className="w-full px-2.5 py-1.5 text-xs bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg tabular-nums"
              />
            </div>
          ))}

          <div className="flex items-center gap-1.5">
            {mode === 'edicion' && (
              <button
                type="button"
                onClick={() => setEditandoLineaId(null)}
                className="px-3 py-1.5 rounded-full text-xs border border-zinc-200 dark:border-[#2E2E35] text-zinc-600 dark:text-zinc-300"
              >
                Cancelar
              </button>
            )}
            <button
              type="button"
              disabled={busy || suma !== l.cantidadTotal}
              onClick={confirmar}
              className="px-3.5 py-1.5 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800 dark:bg-zinc-100 dark:text-zinc-950 dark:hover:bg-white disabled:opacity-40"
            >
              {busy ? 'Guardando...' : 'Confirmar'}
            </button>
          </div>
        </div>

        {!esDos && (
          <div className="text-[11px] tabular-nums">
            <span
              className={
                diff === 0
                  ? 'text-emerald-700 dark:text-emerald-400 font-medium'
                  : diff > 0
                    ? 'text-amber-700 dark:text-amber-400 font-medium'
                    : 'text-red-600 dark:text-red-400 font-medium'
              }
            >
              Asignadas: {suma} / {l.cantidadTotal}
              {diff > 0 ? ` · Faltan ${diff}` : diff < 0 ? ` · Sobran ${Math.abs(diff)}` : ' ✓'}
            </span>
          </div>
        )}
      </div>
    );
  };

  const asignacionesDeposito = pedido.asignaciones.filter((a) => isDepositoRole(a.usuarioRol));
  const esUrgenteDeposito = asignacionesDeposito.some((a) => a.prioridad === 'URGENTE');
  const notasDeposito = Array.from(
    new Set(asignacionesDeposito.map((a) => a.notaEnvio.trim()).filter(Boolean))
  ).join(' · ');
  const requiereRelectura = asignacionesDeposito.some((a) => a.requiereRelectura);
  const totalUnidades = pedido.lineas.reduce((acc, l) => acc + l.cantidadTotal, 0);
  const completoParaDeposito = pedido.faltantesVariante === 0 && pedido.lineas.length > 0;

  const msgReclamoTalles = generarMensajeWhatsApp(
    pedido.cliente,
    pedido.numeroPedido,
    pedido.lineas,
    catalogoCompleto
  );
  const waUrlReclamo = buildWhatsAppUrl(pedido.whatsapp, msgReclamoTalles);

  const msgDespacho = generarMensajeDespachoWhatsApp({
    cliente: pedido.cliente,
    numeroPedido: pedido.numeroPedido,
    transporte: draft.transporte.trim(),
    nroSeguimiento: draft.nroSeguimiento.trim()
  });
  const waUrlDespacho = buildWhatsAppUrl(pedido.whatsapp, msgDespacho);

  const busyHoja = pendingKeys.has(`hoja:${pedido.id}`);
  const busyAcreditar = pendingKeys.has(`acreditar:${pedido.id}`);
  const busyEliminar = pendingKeys.has(`eliminar:${pedido.id}`);
  const busyAjusteArticulos = pendingKeys.has(`articulos:${pedido.id}`);

  const lineasPendientes = pedido.lineas.filter((l) => l.requiereRevision || l.estado !== 'OK');

  const ventasBloqueadoAjuste =
    !esAdmin &&
    (pedido.montoAbonado > 0 ||
      pedido.chequeAcreditado ||
      pedido.estadoEnvio !== 'Pendiente' ||
      pedido.estadoArmado === 'Listo' ||
      asignacionesDeposito.length > 0);
  const puedeAjustarArticulos = !esDeposito && (esAdmin || !ventasBloqueadoAjuste);

  const previewAjuste = useMemo(() => {
    if (!editandoArticulos) return null;
    const previewPedido = recalcularPedido({
      ...pedido,
      lineas: lineasEdicion,
      totalFacturadoOverride: undefined
    });
    const prevTotalCents = toCents(pedido.totalFacturado);
    const nuevoTotalCents = toCents(previewPedido.totalFacturado);
    const prevAbonadoCents = toCents(pedido.montoAbonado);
    const prevChequeCents = toCents(pedido.montoCheque);
    const saldoAFavor = prevAbonadoCents > nuevoTotalCents ? fromCents(prevAbonadoCents - nuevoTotalCents) : 0;

    let diffCheque = 0;
    if (pedido.medioPago === 'Cheque' && !pedido.chequeAcreditado && prevChequeCents > 0) {
      const cubiertoCents = prevAbonadoCents + prevChequeCents;
      if (cubiertoCents > nuevoTotalCents) {
        const nuevoChequeCents = Math.max(0, nuevoTotalCents - Math.min(prevAbonadoCents, nuevoTotalCents));
        diffCheque = fromCents(Math.max(0, prevChequeCents - nuevoChequeCents));
      }
    }

    const unidadesNuevas = lineasEdicion.reduce((acc, l) => acc + l.cantidadTotal, 0);
    const pasoAPendiente =
      pedido.estadoPago === 'Pagado' &&
      nuevoTotalCents > prevAbonadoCents &&
      previewPedido.saldoPendiente > 0.005;

    return {
      previewPedido,
      unidadesNuevas,
      deltaUnidades: unidadesNuevas - totalUnidades,
      deltaTotalCents: nuevoTotalCents - prevTotalCents,
      saldoAFavor,
      diffCheque,
      pasoAPendiente
    };
  }, [editandoArticulos, pedido, lineasEdicion, totalUnidades]);

  const serializarLineasParaComparar = (list: ReadonlyArray<LineaPedido>): string =>
    JSON.stringify(
      list.map((l) => ({
        g: l.idGrupo,
        q: l.cantidadTotal,
        e: l.estado,
        d: l.desglose.map((v) => `${v.codigo}:${v.cantidad}`)
      }))
    );

  const hayCambiosArticulos =
    editandoArticulos &&
    serializarLineasParaComparar(pedido.lineas) !== serializarLineasParaComparar(lineasEdicion);

  const resultadosBusquedaReemplazo = useMemo(() => {
    if (!editandoArticulos || !mostrarBuscadorReemplazo) return [];
    const q = normalizarTexto(busquedaReemplazo);
    return catalogoCompleto
      .filter((prod) => {
        if (!q) return true;
        return (
          normalizarTexto(prod.nombre).includes(q) ||
          normalizarTexto(prod.rubro).includes(q) ||
          prod.variantes.some((v) => normalizarTexto(v.codigo).includes(q))
        );
      })
      .slice(0, 8);
  }, [editandoArticulos, mostrarBuscadorReemplazo, busquedaReemplazo, catalogoCompleto]);

  const handleEntrarAjusteArticulos = (): void => {
    setEditandoLineaId(null);
    setLineasEdicion(
      pedido.lineas.map((l) => ({
        ...l,
        desglose: l.desglose.map((d) => ({ ...d }))
      }))
    );
    setMostrarBuscadorReemplazo(false);
    setBusquedaReemplazo('');
    setReemplazoPendienteProd(null);
    setReemplazoVarsQty({});
    setMostrarAdhocDrawer(false);
    setConfirmQuitarIdGrupo(null);
    setResolverIncidenciaCheck(Boolean(pedido.incidenciaDeposito && esAdmin));
    setConfirmSobreDespachadoCheck(false);
    setEditandoArticulos(true);
  };

  const handleCancelarAjusteArticulos = (): void => {
    setEditandoArticulos(false);
    setLineasEdicion([]);
    setMostrarBuscadorReemplazo(false);
    setReemplazoPendienteProd(null);
    setMostrarAdhocDrawer(false);
    setConfirmQuitarIdGrupo(null);
  };

  const handleQuitarLineaEdicion = (idGrupo: string): void => {
    setLineasEdicion((prev) => prev.filter((l) => l.idGrupo !== idGrupo));
    if (confirmQuitarIdGrupo === idGrupo) setConfirmQuitarIdGrupo(null);
  };

  const handleAjustarCantidadSimpleEdicion = (idGrupo: string, siguienteQty: number): void => {
    if (siguienteQty <= 0) {
      setConfirmQuitarIdGrupo(idGrupo);
      return;
    }
    setConfirmQuitarIdGrupo(null);
    const qty = Math.max(1, Math.trunc(siguienteQty));
    setLineasEdicion((prev) =>
      prev.map((l) => {
        if (l.idGrupo !== idGrupo) return l;
        const nuevoDesglose =
          l.desglose.length === 1 && l.desglose[0]
            ? [{ ...l.desglose[0], cantidad: qty }]
            : l.desglose;
        return {
          ...l,
          cantidadTotal: qty,
          desglose: nuevoDesglose
        };
      })
    );
  };

  const handleAjustarVarianteEdicion = (
    idGrupo: string,
    codigoVar: string,
    siguienteVarQty: number
  ): void => {
    const varQty = Math.max(0, Math.trunc(siguienteVarQty));
    const lineaActual = lineasEdicion.find((l) => l.idGrupo === idGrupo);
    if (!lineaActual) return;
    const prod = catalogoCompleto.find((p) => p.idGrupo === idGrupo);
    const variantesBase = prod ? prod.variantes : lineaActual.desglose;

    const nuevoDesglose = variantesBase.map((v) => {
      const actualVar = lineaActual.desglose.find((d) => d.codigo === v.codigo)?.cantidad ?? 0;
      return {
        codigo: v.codigo,
        etiqueta: v.etiqueta,
        cantidad: v.codigo === codigoVar ? varQty : actualVar
      };
    });

    const nuevaSuma = nuevoDesglose.reduce((acc, d) => acc + d.cantidad, 0);
    if (nuevaSuma <= 0) {
      setConfirmQuitarIdGrupo(idGrupo);
      return;
    }
    setConfirmQuitarIdGrupo(null);
    setLineasEdicion((prev) =>
      prev.map((l) =>
        l.idGrupo === idGrupo
          ? {
              ...l,
              cantidadTotal: nuevaSuma,
              estado: 'OK',
              requiereRevision: false,
              motivoRevision: undefined,
              notaCliente: nuevoDesglose
                .filter((d) => d.cantidad > 0)
                .map((d) => `${d.cantidad} ${d.etiqueta}`)
                .join(' · '),
              desglose: nuevoDesglose
            }
          : l
      )
    );
  };

  const handleSeleccionarReemplazoCatalogo = (prod: ProductoCatalogo): void => {
    if (prod.variantes.length <= 1) {
      const existente = lineasEdicion.find((l) => l.idGrupo === prod.idGrupo);
      if (existente) {
        handleAjustarCantidadSimpleEdicion(prod.idGrupo, existente.cantidadTotal + 1);
      } else {
        const v0 = prod.variantes[0] ?? { codigo: prod.idGrupo, etiqueta: 'Único', aliases: [] };
        setLineasEdicion((prev) => [
          ...prev,
          {
            idLinea: prod.idGrupo,
            idGrupo: prod.idGrupo,
            nombreProducto: prod.nombre,
            rubro: prod.rubro,
            precioUnitarioNeto: prod.precioNeto,
            cantidadTotal: 1,
            notaCliente: '',
            estado: 'OK',
            requiereRevision: false,
            desglose: [{ codigo: v0.codigo, etiqueta: v0.etiqueta, cantidad: 1 }]
          }
        ]);
      }
      setReemplazoPendienteProd(null);
      onNotify(`${prod.nombre} sumado al ajuste.`, 'green');
      return;
    }

    const initMap: Record<string, string> = {};
    for (const v of prod.variantes) initMap[v.codigo] = '0';
    setReemplazoPendienteProd(prod);
    setReemplazoVarsQty(initMap);
  };

  const handleConfirmarReemplazoVariantes = (pendienteDeDefinir = false): void => {
    const prod = reemplazoPendienteProd;
    if (!prod) return;

    if (pendienteDeDefinir) {
      const existente = lineasEdicion.find((l) => l.idGrupo === prod.idGrupo);
      if (existente) {
        handleAjustarCantidadSimpleEdicion(prod.idGrupo, existente.cantidadTotal + 1);
      } else {
        setLineasEdicion((prev) => [...prev, resolverLinea(prod, 1, '')]);
      }
      setReemplazoPendienteProd(null);
      setReemplazoVarsQty({});
      onNotify(`${prod.nombre} agregado (pendiente de talle o color).`, 'amber');
      return;
    }

    const cantidadesPorCod = prod.variantes.map((v) =>
      Math.max(0, Math.trunc(Number(reemplazoVarsQty[v.codigo]) || 0))
    );
    const sumaNueva = cantidadesPorCod.reduce((acc, q) => acc + q, 0);
    if (sumaNueva <= 0) {
      onNotify('Indicá al menos 1 unidad en algún talle o color.', 'amber');
      return;
    }

    const existente = lineasEdicion.find((l) => l.idGrupo === prod.idGrupo);
    if (existente && existente.estado === 'OK' && !existente.requiereRevision) {
      const nuevoDesglose = prod.variantes.map((v, idx) => ({
        codigo: v.codigo,
        etiqueta: v.etiqueta,
        cantidad:
          (existente.desglose.find((d) => d.codigo === v.codigo)?.cantidad ?? 0) +
          (cantidadesPorCod[idx] ?? 0)
      }));
      const totalComb = nuevoDesglose.reduce((acc, d) => acc + d.cantidad, 0);
      setLineasEdicion((prev) =>
        prev.map((l) =>
          l.idGrupo === prod.idGrupo
            ? {
                ...l,
                cantidadTotal: totalComb,
                estado: 'OK',
                requiereRevision: false,
                notaCliente: nuevoDesglose
                  .filter((d) => d.cantidad > 0)
                  .map((d) => `${d.cantidad} ${d.etiqueta}`)
                  .join(' · '),
                desglose: nuevoDesglose
              }
            : l
        )
      );
    } else if (existente) {
      handleAjustarCantidadSimpleEdicion(prod.idGrupo, existente.cantidadTotal + sumaNueva);
    } else {
      const desglose = prod.variantes.map((v, idx) => ({
        codigo: v.codigo,
        etiqueta: v.etiqueta,
        cantidad: cantidadesPorCod[idx] ?? 0
      }));
      setLineasEdicion((prev) => [
        ...prev,
        {
          idLinea: prod.idGrupo,
          idGrupo: prod.idGrupo,
          nombreProducto: prod.nombre,
          rubro: prod.rubro,
          precioUnitarioNeto: prod.precioNeto,
          cantidadTotal: sumaNueva,
          notaCliente: desglose
            .filter((d) => d.cantidad > 0)
            .map((d) => `${d.cantidad} ${d.etiqueta}`)
            .join(' · '),
          estado: 'OK',
          requiereRevision: false,
          desglose
        }
      ]);
    }

    setReemplazoPendienteProd(null);
    setReemplazoVarsQty({});
    onNotify(`${prod.nombre} (${sumaNueva} u.) agregado al ajuste.`, 'green');
  };

  const handleAgregarAdhocEnDrawer = (): void => {
    const nombre = adhocNombre.trim();
    const rubro = adhocRubro.trim().toUpperCase() || 'PILUSOS';
    const precio = round2(parseFloat(adhocPrecio.replace(',', '.')) || 0);
    const qty = Math.max(0, Math.trunc(Number(adhocCantidad) || 0));
    if (!nombre) {
      onNotify('Ingresá el nombre del artículo fuera de planilla.', 'amber');
      return;
    }
    if (qty <= 0) {
      onNotify('Ingresá una cantidad mayor a 0.', 'amber');
      return;
    }
    const idAdhoc = `ADHOC-${Date.now().toString(36).toUpperCase()}`;
    setLineasEdicion((prev) => [
      ...prev,
      {
        idLinea: idAdhoc,
        idGrupo: idAdhoc,
        nombreProducto: nombre,
        rubro,
        precioUnitarioNeto: Math.max(0, precio),
        cantidadTotal: qty,
        notaCliente: '',
        estado: 'OK',
        requiereRevision: false,
        desglose: [{ codigo: idAdhoc, etiqueta: 'Único', cantidad: qty }]
      }
    ]);
    setAdhocNombre('');
    setAdhocPrecio('');
    setAdhocCantidad('1');
    setMostrarAdhocDrawer(false);
    onNotify(`${nombre} (${qty} u.) agregado al ajuste.`, 'green');
  };

  const handleGuardarAjusteArticulos = async (): Promise<void> => {
    if (busyAjusteArticulos || lineasEdicion.length === 0 || !hayCambiosArticulos) return;
    if (pedido.estadoEnvio !== 'Pendiente' && !confirmSobreDespachadoCheck) {
      onNotify('Confirmá el ajuste sobre el pedido ya despachado antes de guardar.', 'amber');
      return;
    }
    const resumen = await onActualizarArticulos({
      pedidoId: pedido.id,
      expectedVersion: pedido.version,
      lineas: lineasEdicion,
      resolverIncidencia: resolverIncidenciaCheck,
      confirmarSobreDespachado: pedido.estadoEnvio !== 'Pendiente' ? confirmSobreDespachadoCheck : undefined
    });
    if (resumen) {
      setEditandoArticulos(false);
      setLineasEdicion([]);
      setMostrarBuscadorReemplazo(false);
      setReemplazoPendienteProd(null);
      setConfirmQuitarIdGrupo(null);
      setUltimoAjusteGuardado(resumen);
    }
  };

  const msgAjusteWhatsapp = ultimoAjusteGuardado
    ? generarMensajeAjusteStockWhatsApp(pedido, ultimoAjusteGuardado)
    : '';
  const waUrlAjusteWhatsapp = msgAjusteWhatsapp ? buildWhatsAppUrl(pedido.whatsapp, msgAjusteWhatsapp) : null;

  const renderHojaDeposito = (modoVistaPrevia: boolean): React.ReactNode => (
    <div
      className={
        modoVistaPrevia
          ? 'space-y-4 bg-white text-[#111111] border border-zinc-200 rounded-2xl p-5'
          : 'space-y-4 bg-white text-[#111111]'
      }
    >
      <div className="flex items-start justify-between border-b border-zinc-200 pb-3">
        <div>
          <MibiLogo width={76} variant="pink" />
          <div className="text-xs font-semibold text-[#111111] mt-2 flex items-center gap-2">
            <span>HOJA PARA DEPÓSITO — PEDIDO #{pedido.numeroPedido}</span>
            {esUrgenteDeposito && (
              <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-red-100 text-red-700">
                URGENTE
              </span>
            )}
          </div>
          <div className="text-[11px] text-zinc-600">
            Cliente: {pedido.cliente} · Transporte: {pedido.transporte || 'A definir'}
            {pedido.despachoPrevisto ? ` · Despacho: ${formatShortDate(pedido.despachoPrevisto)}` : ''}
          </div>
          {notasDeposito && (
            <div className="text-[11px] text-zinc-700 mt-1">
              <strong>Nota operativa:</strong> {notasDeposito}
            </div>
          )}
        </div>
        {modoVistaPrevia && (
          <button
            type="button"
            onClick={() => window.print()}
            className="no-print px-3 py-1.5 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800"
          >
            Imprimir ahora
          </button>
        )}
      </div>

      <table className="w-full text-xs border-collapse text-[#111111]">
        <thead>
          <tr className="border-b border-zinc-300 text-zinc-600">
            <th className="py-2 text-left">Código</th>
            <th className="py-2 text-left">Producto</th>
            <th className="py-2 text-left">Talle / Color</th>
            <th className="py-2 text-right">Cantidad pedida</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-zinc-200/70 tabular-nums">
          {pedido.lineas.flatMap((l) =>
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
          <tr className="border-t-2 border-zinc-300 font-semibold">
            <td colSpan={3} className="py-2.5 text-right text-zinc-600">
              TOTAL UNIDADES
            </td>
            <td className="py-2.5 text-right tabular-nums">{totalUnidades}</td>
          </tr>
        </tbody>
      </table>
      <p className="mibi-print-end-block text-[10px] text-zinc-500 pt-2 border-t border-zinc-200">
        Mibi® · Ciudad de la Paz 2467, Belgrano, CABA · WhatsApp 11-5606-9384 · contacto@mibi.com.ar
      </p>
    </div>
  );

  return (
    <>
    <div
      className="fixed inset-0 z-50 flex justify-end bg-black/25 dark:bg-black/60 backdrop-blur-[1px]"
      onClick={onClose}
    >
      <datalist id={datalistId}>
        {transportesSugeridos.map((t) => (
          <option key={t} value={t} />
        ))}
      </datalist>

      <aside
        onClick={(e) => e.stopPropagation()}
        className="w-full max-w-xl bg-white dark:bg-[#141417] h-full overflow-y-auto border-l border-[#EAEAEA] dark:border-[#26262B] shadow-2xl flex flex-col"
      >
        <div className="sticky top-0 z-10 bg-white/95 dark:bg-[#141417]/90 backdrop-blur border-b border-[#EAEAEA] dark:border-[#26262B] px-6 py-4 flex items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={() =>
                  onCopyText(`num-${pedido.id}`, pedido.numeroPedido, `Número #${pedido.numeroPedido} copiado.`)
                }
                title="Clic para copiar número de pedido"
                className="text-base font-semibold text-[#111111] dark:text-[#F4F4F5] tabular-nums hover:underline inline-flex items-center gap-1"
              >
                <span>Pedido #{pedido.numeroPedido}</span>
                {copiadoId === `num-${pedido.id}` && (
                  <Check className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400" strokeWidth={2} />
                )}
              </button>
              {pedido.faltantesVariante > 0 ? (
                <StatusPill tone="red">Falta talle o color</StatusPill>
              ) : (
                <StatusPill tone={pedido.estadoArmado === 'Listo' ? 'green' : 'neutral'}>
                  {pedido.estadoArmado}
                </StatusPill>
              )}
              {esUrgenteDeposito && <StatusPill tone="red">Urgente</StatusPill>}
              {esDeposito && requiereRelectura && <StatusPill tone="amber">Actualizado</StatusPill>}
            </div>
            <p className="text-xs text-zinc-500 dark:text-zinc-400 mt-0.5">
              {pedido.cliente} · {pedido.vendedor}
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-2 rounded-full text-zinc-400 hover:text-zinc-900 hover:bg-zinc-100 dark:hover:text-[#F4F4F5] dark:hover:bg-[#1B1B1F]"
            aria-label="Cerrar panel"
          >
            <X className="w-4 h-4" strokeWidth={1.5} />
          </button>
        </div>

        <div className="p-6 space-y-6 flex-1">
          <div className="flex items-center justify-between gap-2">
            <div className="flex flex-wrap items-center gap-2">
              {completoParaDeposito && !mostrarHojaImprimible && (
                <>
                  <button
                    type="button"
                    disabled={busyHoja}
                    onClick={() => onExportarHojaDeposito(pedido)}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800 dark:bg-zinc-100 dark:text-zinc-950 dark:hover:bg-white transition-colors disabled:opacity-40"
                  >
                    <Download className="w-3.5 h-3.5" strokeWidth={1.5} />
                    {busyHoja ? 'Descargando...' : 'Hoja depósito (.xlsx)'}
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      document.body.classList.add('mibi-printing-hoja');
                      setMostrarHojaImprimible(true);
                      window.setTimeout(() => window.print(), 100);
                    }}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 dark:bg-[#1B1B1F] dark:border-[#2E2E35] dark:text-zinc-200 dark:hover:bg-[#24242B] transition-colors"
                  >
                    <Printer className="w-3.5 h-3.5" strokeWidth={1.5} />
                    Imprimir
                  </button>
                </>
              )}

              {mostrarHojaImprimible && (
                <button
                  type="button"
                  onClick={() => setMostrarHojaImprimible(false)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 dark:bg-[#1B1B1F] dark:border-[#2E2E35] dark:text-zinc-200 dark:hover:bg-[#24242B] transition-colors"
                >
                  Volver a gestión
                </button>
              )}
            </div>

            {!esDeposito && !mostrarHojaImprimible && (
              <div ref={moreMenuRef} className="relative ml-auto">
                <button
                  type="button"
                  onClick={() => setMenuMasAbierto((prev) => !prev)}
                  aria-label="Más acciones"
                  className="p-2 rounded-full border border-zinc-200 text-zinc-600 hover:bg-zinc-50 dark:bg-[#1B1B1F] dark:border-[#2E2E35] dark:text-zinc-300 dark:hover:bg-[#24242B] transition-colors"
                >
                  <MoreHorizontal className="w-4 h-4" strokeWidth={1.5} />
                </button>

                {menuMasAbierto && (
                  <div className="absolute right-0 mt-1.5 w-52 bg-white dark:bg-[#18181D] rounded-2xl border border-[#EAEAEA] dark:border-[#2E2E36] shadow-[0_12px_36px_rgba(0,0,0,0.12)] dark:shadow-[0_12px_36px_rgba(0,0,0,0.55)] p-1.5 z-30 text-xs">
                    <button
                      type="button"
                      onClick={() => {
                        setMenuMasAbierto(false);
                        onDuplicarComoNuevo(pedido);
                      }}
                      className="w-full text-left px-3 py-2 rounded-xl flex items-center gap-2 text-[#111111] dark:text-[#F4F4F5] hover:bg-zinc-100 dark:hover:bg-[#24242B]"
                    >
                      <Copy className="w-3.5 h-3.5 text-zinc-500 dark:text-zinc-400" strokeWidth={1.5} />
                      <span>Duplicar como nuevo</span>
                    </button>

                    {esAdmin && (
                      <button
                        type="button"
                        onClick={() => {
                          setMenuMasAbierto(false);
                          setConfirmEliminar(true);
                        }}
                        className="w-full text-left px-3 py-2 rounded-xl flex items-center gap-2 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-500/10"
                      >
                        <Trash2 className="w-3.5 h-3.5" strokeWidth={1.5} />
                        <span>Eliminar pedido</span>
                      </button>
                    )}
                  </div>
                )}
              </div>
            )}
          </div>

          {confirmEliminar && esAdmin && !mostrarHojaImprimible && (
            <div className="bg-red-50/80 border border-red-200 dark:bg-red-500/10 dark:border-red-500/25 rounded-2xl px-4 py-3 flex items-center justify-between gap-3 text-xs">
              <span className="text-red-800 dark:text-red-200 font-medium">
                ¿Eliminar el pedido #{pedido.numeroPedido}?
              </span>
              <div className="flex items-center gap-2 shrink-0">
                <button
                  type="button"
                  onClick={() => setConfirmEliminar(false)}
                  className="px-3 py-1 rounded-full bg-white border border-zinc-200 text-zinc-700 dark:bg-[#1B1B1F] dark:border-[#2E2E35] dark:text-zinc-200"
                >
                  Cancelar
                </button>
                <button
                  type="button"
                  disabled={busyEliminar}
                  onClick={() => onEliminarIndividual(pedido)}
                  className="px-3.5 py-1 rounded-full bg-red-600 hover:bg-red-700 text-white font-medium disabled:opacity-40"
                >
                  {busyEliminar ? 'Eliminando...' : 'Confirmar'}
                </button>
              </div>
            </div>
          )}

          {mostrarHojaImprimible ? (
            renderHojaDeposito(true)
          ) : (
            <>
              {esDeposito && (notasDeposito || esUrgenteDeposito || requiereRelectura) && (
                <div
                  className={`rounded-2xl border p-4 space-y-1.5 text-xs ${
                    esUrgenteDeposito
                      ? 'bg-red-50/70 border-red-200 text-red-900 dark:bg-red-500/10 dark:border-red-500/25 dark:text-red-200'
                      : 'bg-amber-50/70 border-amber-200 text-amber-900 dark:bg-amber-500/10 dark:border-amber-500/25 dark:text-amber-200'
                  }`}
                >
                  <div className="font-semibold flex items-center justify-between gap-2">
                    <span>
                      {esUrgenteDeposito ? 'Prioridad Urgente — Producción' : 'Instrucción de Administración'}
                    </span>
                    {pedido.despachoPrevisto && (
                      <span className="text-[11px] tabular-nums">
                        Despacho previsto: {formatShortDate(pedido.despachoPrevisto)}
                      </span>
                    )}
                  </div>
                  {notasDeposito && <p className="text-xs">{notasDeposito}</p>}
                  {requiereRelectura && (
                    <p className="text-[11px] font-medium">
                      Administración actualizó los artículos de este pedido. Revisá el detalle antes de embalar.
                    </p>
                  )}
                </div>
              )}

              {pedido.incidenciaDeposito && (
                <div className="bg-amber-50/80 border border-amber-200 dark:bg-amber-500/10 dark:border-amber-500/25 rounded-2xl p-4 space-y-2 text-xs">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-amber-900 dark:text-amber-200 flex items-center gap-1.5">
                      <AlertTriangle className="w-4 h-4 text-amber-600 dark:text-amber-400" strokeWidth={1.75} />
                      Aviso de Depósito {pedido.incidenciaPor ? `(${pedido.incidenciaPor})` : ''}
                    </span>
                    {esAdmin && (
                      <button
                        type="button"
                        disabled={pendingKeys.has(`resolverInc:${pedido.id}`)}
                        onClick={() => void onResolverIncidencia(pedido.id)}
                        className="px-3 py-1 rounded-full text-[11px] font-medium bg-white border border-amber-300 text-amber-900 hover:bg-amber-100 dark:bg-[#1B1B1F] dark:border-amber-500/35 dark:text-amber-200 dark:hover:bg-amber-500/20 disabled:opacity-40"
                      >
                        Marcar resuelto
                      </button>
                    )}
                  </div>
                  <p className="text-amber-900 dark:text-amber-200/90">{pedido.incidenciaDeposito}</p>
                </div>
              )}

              {!esDeposito && pedido.faltantesVariante > 0 && (
                <div
                  ref={tallesSectionRef}
                  className="bg-red-50/70 border border-red-200 dark:bg-red-500/10 dark:border-red-500/25 rounded-2xl p-4 space-y-3"
                >
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <span className="text-xs font-semibold text-red-800 dark:text-red-300">
                      Falta definir talle o color ({pedido.faltantesVariante})
                    </span>

                    <div className="inline-flex items-center gap-1.5">
                      {waUrlReclamo ? (
                        <>
                          <a
                            href={waUrlReclamo}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors"
                          >
                            <MessageCircle className="w-3 h-3" strokeWidth={1.75} />
                            Enviar por WhatsApp
                          </a>
                          <button
                            type="button"
                            title="Copiar mensaje"
                            aria-label="Copiar mensaje de reclamo"
                            onClick={() =>
                              onCopyText(`wsp-${pedido.id}`, msgReclamoTalles, 'Mensaje copiado para WhatsApp.')
                            }
                            className="p-1.5 rounded-full bg-white border border-red-200 text-red-700 hover:bg-red-50 dark:bg-[#1B1B1F] dark:border-red-500/30 dark:text-red-300"
                          >
                            {copiadoId === `wsp-${pedido.id}` ? (
                              <Check className="w-3 h-3 text-emerald-600 dark:text-emerald-400" strokeWidth={1.75} />
                            ) : (
                              <Copy className="w-3 h-3" strokeWidth={1.5} />
                            )}
                          </button>
                        </>
                      ) : (
                        <button
                          type="button"
                          onClick={() =>
                            onCopyText(
                              `wsp-${pedido.id}`,
                              msgReclamoTalles,
                              'Reclamo para WhatsApp copiado.'
                            )
                          }
                          className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-medium bg-white border border-red-200 text-red-700 hover:bg-red-50 dark:bg-[#1B1B1F] dark:border-red-500/30 dark:text-red-300 dark:hover:bg-red-500/15"
                        >
                          {copiadoId === `wsp-${pedido.id}` ? (
                            <Check className="w-3 h-3 text-emerald-600 dark:text-emerald-400" strokeWidth={1.75} />
                          ) : (
                            <MessageCircle className="w-3 h-3" strokeWidth={1.5} />
                          )}
                          {copiadoId === `wsp-${pedido.id}` ? 'Copiado' : 'Copiar para WhatsApp'}
                        </button>
                      )}
                    </div>
                  </div>

                  {lineasPendientes.map((l, idx) => renderVariantEditor(l, 'pendiente', idx === 0))}
                </div>
              )}

              {esAdmin && (
                <div
                  ref={derivacionSectionRef}
                  className="space-y-3 bg-[#FAFAFA] dark:bg-[#1B1B1F] rounded-2xl border border-zinc-200/80 dark:border-[#26262B] p-4 text-xs"
                >
                  <div className="flex items-center justify-between gap-2">
                    <h3 className="text-xs font-semibold text-zinc-600 dark:text-zinc-300 uppercase tracking-wider flex items-center gap-1.5">
                      <UserCheck className="w-3.5 h-3.5" strokeWidth={1.5} />
                      Derivación a cuentas
                    </h3>
                    <button
                      type="button"
                      onClick={() => {
                        setMostrarEnviar((prev) => !prev);
                        setRequiereConfirmSinPago(false);
                      }}
                      className="px-3 py-1 rounded-full text-[11px] font-medium bg-zinc-900 text-white hover:bg-zinc-800 dark:bg-zinc-100 dark:text-zinc-950 dark:hover:bg-white"
                    >
                      {mostrarEnviar ? 'Cerrar envío' : '+ Enviar a cuenta'}
                    </button>
                  </div>

                  {pedido.asignaciones.length === 0 ? (
                    <p className="text-[11px] text-zinc-500 dark:text-zinc-400">
                      Solo en Administración · Todavía no fue enviado a Depósito ni a Ventas.
                    </p>
                  ) : (
                    <div className="space-y-1.5">
                      {pedido.asignaciones.map((a) => (
                        <div
                          key={a.id}
                          className="flex items-center justify-between gap-2 bg-white dark:bg-[#141417] border border-zinc-200/80 dark:border-[#26262B] rounded-xl px-3 py-2"
                        >
                          <div className="min-w-0">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <span className="font-semibold text-[#111111] dark:text-[#F4F4F5]">{a.usuarioNombre}</span>
                              {a.prioridad === 'URGENTE' && <StatusPill tone="red">Urgente</StatusPill>}
                              {a.requiereRelectura ? (
                                <StatusPill tone="amber">Actualizado</StatusPill>
                              ) : a.vistoAt ? (
                                <StatusPill tone="green">Visto {formatHoraCorta(a.vistoAt)}</StatusPill>
                              ) : (
                                <StatusPill tone="neutral">Sin leer</StatusPill>
                              )}
                            </div>
                            {a.notaEnvio && (
                              <div className="text-[11px] text-zinc-500 dark:text-zinc-400 truncate mt-0.5">{a.notaEnvio}</div>
                            )}
                          </div>
                          <button
                            type="button"
                            disabled={pendingKeys.has(`retirar:${pedido.id}:${a.usuarioId}`)}
                            onClick={() => void onRetirarAsignacion(pedido.id, a.usuarioId)}
                            className="p-1 text-zinc-400 hover:text-red-600 dark:hover:text-red-400 disabled:opacity-40"
                            title={`Retirar de ${a.usuarioNombre}`}
                            aria-label={`Retirar asignación de ${a.usuarioNombre}`}
                          >
                            <X className="w-3.5 h-3.5" strokeWidth={1.75} />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}

                  {mostrarEnviar && (
                    <div className="pt-2 border-t border-zinc-200/70 dark:border-[#26262B] space-y-2.5">
                      <div>
                        <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Cuenta destino</label>
                        <select
                          value={cuentaDestinoId}
                          onChange={(e) => {
                            setCuentaDestinoId(e.target.value);
                            setRequiereConfirmSinPago(false);
                          }}
                          className="w-full px-3 py-2 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl"
                        >
                          <option value="">Elegir cuenta...</option>
                          {cuentasDeposito.length > 0 && (
                            <optgroup label="Depósito / Producción">
                              {cuentasDeposito.map((c) => (
                                <option key={c.id} value={c.id}>
                                  {c.nombre} (@{c.username})
                                </option>
                              ))}
                            </optgroup>
                          )}
                          {cuentasVentas.length > 0 && (
                            <optgroup label="Ventas">
                              {cuentasVentas.map((c) => (
                                <option key={c.id} value={c.id}>
                                  {c.nombre} (@{c.username})
                                </option>
                              ))}
                            </optgroup>
                          )}
                        </select>
                      </div>

                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Prioridad</label>
                          <select
                            value={prioridadEnvio}
                            onChange={(e) => setPrioridadEnvio(e.target.value as PrioridadAsignacion)}
                            className="w-full px-3 py-2 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl"
                          >
                            <option value="NORMAL">Normal</option>
                            <option value="URGENTE">Urgente</option>
                          </select>
                        </div>
                        <div>
                          <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Nota breve (opc.)</label>
                          <input
                            type="text"
                            maxLength={300}
                            value={notaEnvio}
                            onChange={(e) => setNotaEnvio(e.target.value)}
                            placeholder="Ej. Sale transporte 15hs"
                            className="w-full px-3 py-2 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                          />
                        </div>
                      </div>

                      {requiereConfirmSinPago ? (
                        <div className="bg-amber-50 border border-amber-200 dark:bg-amber-500/10 dark:border-amber-500/25 rounded-xl p-3 space-y-2">
                          <p className="text-[11px] text-amber-900 dark:text-amber-200 font-medium">
                            Este pedido figura sin pago registrado. Confirmá si querés liberarlo igual a Depósito.
                          </p>
                          <div className="flex gap-2">
                            <button
                              type="button"
                              onClick={() => setRequiereConfirmSinPago(false)}
                              className="flex-1 py-1.5 rounded-full bg-white border border-zinc-200 text-zinc-700 dark:bg-[#141417] dark:border-[#2E2E35] dark:text-zinc-200"
                            >
                              Cancelar
                            </button>
                            <button
                              type="button"
                              disabled={enviandoCuenta}
                              onClick={() => void handleConfirmarEnvioCuenta(true)}
                              className="flex-1 py-1.5 rounded-full bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-950 font-medium disabled:opacity-40"
                            >
                              Liberar igual
                            </button>
                          </div>
                        </div>
                      ) : (
                        <button
                          type="button"
                          disabled={enviandoCuenta || !cuentaDestinoId}
                          onClick={() => void handleConfirmarEnvioCuenta(false)}
                          className="w-full py-2 rounded-full bg-zinc-900 hover:bg-zinc-800 text-white dark:bg-zinc-100 dark:hover:bg-white dark:text-zinc-950 font-medium disabled:opacity-40"
                        >
                          {enviandoCuenta ? 'Enviando...' : 'Confirmar envío'}
                        </button>
                      )}
                    </div>
                  )}
                </div>
              )}

              {!esDeposito && (
                <div className="space-y-3">
                  <div className="flex items-center justify-between gap-2">
                    <h3 className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 uppercase tracking-wider">
                      Datos del cliente
                    </h3>
                    {renderSaveBadge('cliente')}
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                    <div>
                      <label htmlFor={`drawer-cliente-${pedido.id}`} className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                        Cliente / Comercio
                      </label>
                      <input
                        id={`drawer-cliente-${pedido.id}`}
                        type="text"
                        maxLength={200}
                        value={draft.cliente}
                        onChange={(e) => setDraft((prev) => ({ ...prev, cliente: e.target.value }))}
                        onBlur={commitCliente}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') e.currentTarget.blur();
                        }}
                        placeholder="Nombre del cliente o local"
                        className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
                      />
                    </div>
                    <div>
                      <label htmlFor={`drawer-whatsapp-${pedido.id}`} className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                        WhatsApp
                      </label>
                      <input
                        id={`drawer-whatsapp-${pedido.id}`}
                        type="text"
                        maxLength={80}
                        value={draft.whatsapp}
                        onChange={(e) => setDraft((prev) => ({ ...prev, whatsapp: e.target.value }))}
                        onBlur={commitWhatsapp}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') e.currentTarget.blur();
                        }}
                        placeholder="Ej. 11-5555-1234"
                        className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
                      />
                    </div>
                  </div>
                </div>
              )}

              {!esDeposito && (
                <div className="pt-5 border-t border-zinc-200/70 dark:border-[#26262B] space-y-3">
                  <div className="flex items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <h3 className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 uppercase tracking-wider">
                        Pagos y liquidación
                      </h3>
                      {renderSaveBadge('pagos')}
                    </div>
                    {pedido.estadoPago === 'Pagado' ? (
                      <StatusPill tone="green">Pagado</StatusPill>
                    ) : pedido.estadoPago === 'Pendiente' ? (
                      <StatusPill tone="amber">Pendiente</StatusPill>
                    ) : (
                      <StatusPill tone="red">Falta pago</StatusPill>
                    )}
                  </div>

                  <div className="bg-[#FAFAFA] dark:bg-[#1B1B1F] rounded-2xl border border-zinc-200/80 dark:border-[#26262B] p-4 space-y-1.5 text-xs tabular-nums">
                    <div className="flex justify-between text-zinc-500 dark:text-zinc-400">
                      <span>Subtotal Neto</span>
                      <span>{formatARS(pedido.subtotalNeto)}</span>
                    </div>
                    {pedido.descuentoPorc > 0 && (
                      <div className="flex justify-between text-zinc-500 dark:text-zinc-400">
                        <span>Descuento ({pedido.descuentoPorc}%)</span>
                        <span>-{formatARS(pedido.montoDescuento)}</span>
                      </div>
                    )}
                    <div className="flex justify-between text-zinc-500 dark:text-zinc-400">
                      <span>Base Imponible (Total B)</span>
                      <span>{formatARS(pedido.baseImponible)}</span>
                    </div>
                    <div className="flex justify-between text-zinc-500 dark:text-zinc-400">
                      <span>IVA 21%</span>
                      <span>{formatARS(pedido.iva21)}</span>
                    </div>
                    <div className="flex justify-between font-semibold text-sm text-[#111111] dark:text-[#F4F4F5] pt-1.5 border-t border-zinc-200/70 dark:border-[#26262B]">
                      <span>Total Facturado</span>
                      <span>{formatARS(pedido.totalFacturado)}</span>
                    </div>
                    <div className="flex justify-between text-emerald-700 dark:text-emerald-400 font-medium">
                      <span>Cobrado en mano</span>
                      <span>{formatARS(pedido.montoAbonado)}</span>
                    </div>
                    <div className="flex justify-between text-zinc-700 dark:text-zinc-300 font-medium">
                      <span>Saldo pendiente</span>
                      <span>{formatARS(pedido.saldoPendiente)}</span>
                    </div>
                  </div>

                  {esAdmin ? (
                    pedido.estadoPago === 'Pagado' && !editandoPago ? (
                      <div className="flex items-center justify-between gap-2 bg-emerald-50/70 border border-emerald-200/80 dark:bg-emerald-500/10 dark:border-emerald-500/25 rounded-xl px-3.5 py-2.5 text-xs">
                        <span className="font-medium text-emerald-800 dark:text-emerald-300">
                          Pagado al 100% · Cobrado en mano ({formatARS(pedido.montoAbonado)})
                        </span>
                        <button
                          type="button"
                          onClick={() => setEditandoPago(true)}
                          className="text-[11px] font-medium text-emerald-900 dark:text-emerald-200 underline underline-offset-2 shrink-0"
                        >
                          Editar pago
                        </button>
                      </div>
                    ) : pedido.medioPago === 'Cheque' && !pedido.chequeAcreditado && !cambiandoFormaCheque ? (
                      <div className="space-y-3 text-xs">
                        <button
                          type="button"
                          disabled={busyAcreditar}
                          onClick={() => void onConfirmarAcreditacion(pedido)}
                          className="w-full py-2.5 rounded-full text-xs font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors disabled:opacity-40"
                        >
                          {busyAcreditar
                            ? 'Acreditando...'
                            : `Acreditar cheque (${formatARSCompacto(pedido.montoCheque || pedido.saldoPendiente)}) → Pasar a Pagado`}
                        </button>

                        <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                          <div>
                            <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                              Recepción cheque
                            </label>
                            <input
                              type="date"
                              value={draft.fechaRecepcionCheque}
                              onChange={(e) => {
                                const rec = e.target.value;
                                setDraft((prev) => ({
                                  ...prev,
                                  fechaRecepcionCheque: rec,
                                  ...(rec ? { fechaAcreditacionCheque: addDaysISO(rec, 30) } : {})
                                }));
                              }}
                              onBlur={() => {
                                const f = draft.fechaRecepcionCheque;
                                if (f && f !== pedido.fechaRecepcionCheque) {
                                  void patchConEstado(
                                    'pagos',
                                    { fechaRecepcionCheque: f },
                                    `Actualizó recepción de cheque a ${formatShortDate(f)} (+30d)`
                                  );
                                }
                              }}
                              className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums"
                            />
                          </div>
                          <div>
                            <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                              Vencimiento (+30d)
                            </label>
                            <input
                              type="date"
                              value={draft.fechaAcreditacionCheque}
                              onChange={(e) =>
                                setDraft((prev) => ({ ...prev, fechaAcreditacionCheque: e.target.value }))
                              }
                              onBlur={() => {
                                const f = draft.fechaAcreditacionCheque;
                                if (f && f !== pedido.fechaAcreditacionCheque) {
                                  void patchConEstado(
                                    'pagos',
                                    { fechaAcreditacionCheque: f },
                                    `Actualizó vencimiento de cheque a ${formatShortDate(f)}`
                                  );
                                }
                              }}
                              className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums"
                            />
                          </div>
                          <div>
                            <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                              Detalle cheque
                            </label>
                            <input
                              type="text"
                              value={draft.detalleCheque}
                              onChange={(e) => setDraft((prev) => ({ ...prev, detalleCheque: e.target.value }))}
                              onBlur={() => {
                                const d = draft.detalleCheque.trim();
                                if (d !== pedido.detalleCheque) {
                                  void patchConEstado(
                                    'pagos',
                                    { detalleCheque: d },
                                    'Actualizó detalle de cheque'
                                  );
                                }
                              }}
                              placeholder="Banco / N.º"
                              className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                            />
                          </div>
                        </div>

                        <div className="flex justify-end">
                          <button
                            type="button"
                            onClick={() => setCambiandoFormaCheque(true)}
                            className="text-[11px] text-zinc-500 hover:text-[#111111] dark:text-zinc-400 dark:hover:text-[#F4F4F5] underline underline-offset-2"
                          >
                            Cambiar forma de pago
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-3 text-xs">
                        {editandoPago && pedido.estadoPago === 'Pagado' && (
                          <div className="flex justify-end">
                            <button
                              type="button"
                              onClick={() => setEditandoPago(false)}
                              className="text-[11px] text-zinc-500 hover:text-[#111111] dark:text-zinc-400 dark:hover:text-[#F4F4F5] underline underline-offset-2"
                            >
                              Ocultar edición de pago
                            </button>
                          </div>
                        )}

                        <div className="grid grid-cols-3 p-1 rounded-xl bg-zinc-100 dark:bg-[#1B1B1F] border border-zinc-200/70 dark:border-[#26262B] gap-1">
                          {(['Sin pagar', 'Transferencia / Efectivo', 'Cheque'] as const).map((m) => (
                            <button
                              key={m}
                              type="button"
                              onClick={() => {
                                setCambiandoFormaCheque(false);
                                if (m === 'Transferencia / Efectivo') {
                                  void patchConEstado(
                                    'pagos',
                                    {
                                      medioPago: m,
                                      montoAbonado: pedido.totalFacturado,
                                      montoCheque: 0,
                                      chequeAcreditado: true
                                    },
                                    'Marcó pago completo (Transferencia / Efectivo)'
                                  ).then((ok) => {
                                    if (ok) setEditandoPago(false);
                                  });
                                } else if (m === 'Cheque') {
                                  const hoy = getTodayISO();
                                  void patchConEstado(
                                    'pagos',
                                    {
                                      medioPago: 'Cheque',
                                      chequeAcreditado: false,
                                      montoCheque: round2(pedido.totalFacturado - pedido.montoAbonado),
                                      fechaRecepcionCheque: pedido.fechaRecepcionCheque || hoy,
                                      fechaAcreditacionCheque:
                                        pedido.fechaAcreditacionCheque || addDaysISO(hoy, 30)
                                    },
                                    'Registró pago con Cheque (30 días)'
                                  );
                                } else {
                                  void patchConEstado(
                                    'pagos',
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
                              className={`min-h-[36px] px-2 rounded-lg text-[11px] font-medium transition-colors ${
                                pedido.medioPago === m
                                  ? 'bg-white text-[#111111] dark:bg-[#2A2A32] dark:text-white shadow-sm'
                                  : 'text-zinc-600 hover:text-[#111111] dark:text-zinc-400 dark:hover:text-[#F4F4F5]'
                              }`}
                            >
                              {m === 'Transferencia / Efectivo' ? 'Pago / Seña' : m}
                            </button>
                          ))}
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                          <div>
                            <div className="flex items-center justify-between gap-1 mb-1">
                              <label className="text-[11px] text-zinc-500 dark:text-zinc-400">
                                Monto abonado ($)
                              </label>
                              {pedido.saldoPendiente > 0.01 && (
                                <div className="flex items-center gap-1">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setDraft((prev) => ({
                                        ...prev,
                                        montoAbonado: String(pedido.totalFacturado)
                                      }));
                                      void patchConEstado(
                                        'pagos',
                                        {
                                          medioPago: 'Transferencia / Efectivo',
                                          chequeAcreditado: true,
                                          montoAbonado: pedido.totalFacturado,
                                          montoCheque: 0
                                        },
                                        `Saldó resto (${formatARS(pedido.totalFacturado)})`
                                      );
                                    }}
                                    className="px-2 py-0.5 rounded-full text-[10px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200/70 hover:bg-emerald-100 dark:bg-emerald-500/15 dark:text-emerald-300 dark:border-emerald-500/30"
                                  >
                                    Saldar resto
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const mitad = round2(pedido.totalFacturado / 2);
                                      setDraft((prev) => ({ ...prev, montoAbonado: String(mitad) }));
                                      void patchConEstado(
                                        'pagos',
                                        {
                                          ...(pedido.medioPago === 'Sin pagar'
                                            ? { medioPago: 'Transferencia / Efectivo' }
                                            : {}),
                                          montoAbonado: mitad
                                        },
                                        `Registró seña 50% (${formatARS(mitad)})`
                                      );
                                    }}
                                    className="px-2 py-0.5 rounded-full text-[10px] font-medium bg-zinc-100 text-zinc-700 border border-zinc-200/80 hover:bg-zinc-200/80 dark:bg-[#141417] dark:text-zinc-300 dark:border-[#2E2E35]"
                                  >
                                    50%
                                  </button>
                                </div>
                              )}
                            </div>
                            <input
                              type="number"
                              min={0}
                              max={pedido.totalFacturado}
                              step="0.01"
                              value={draft.montoAbonado}
                              onWheel={(e) => e.currentTarget.blur()}
                              onChange={(e) => setDraft((prev) => ({ ...prev, montoAbonado: e.target.value }))}
                              onBlur={() => {
                                const val = Math.max(
                                  0,
                                  Math.min(pedido.totalFacturado, parseFloat(draft.montoAbonado) || 0)
                                );
                                if (Math.abs(val - pedido.montoAbonado) > 0.005) {
                                  void patchConEstado(
                                    'pagos',
                                    {
                                      montoAbonado: round2(val),
                                      ...(val > 0 && pedido.medioPago === 'Sin pagar'
                                        ? { medioPago: 'Transferencia / Efectivo' }
                                        : {})
                                    },
                                    `Actualizó monto abonado a ${formatARS(val)}`
                                  );
                                }
                              }}
                              className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums"
                            />
                          </div>

                          <div>
                            <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                              Descuento (%)
                            </label>
                            <input
                              type="number"
                              min={0}
                              max={100}
                              step="0.5"
                              disabled={pedido.lineas.length === 0}
                              value={draft.descuentoPorc}
                              onWheel={(e) => e.currentTarget.blur()}
                              onChange={(e) => setDraft((prev) => ({ ...prev, descuentoPorc: e.target.value }))}
                              onBlur={() => {
                                if (pedido.lineas.length === 0) {
                                  onNotify('Los pedidos históricos sin detalle mantienen su total importado.', 'amber');
                                  return;
                                }
                                const val = Math.max(0, Math.min(100, parseFloat(draft.descuentoPorc) || 0));
                                if (val !== pedido.descuentoPorc) {
                                  const prevAbonado = pedido.montoAbonado;
                                  void patchConEstado(
                                    'pagos',
                                    { descuentoPorc: val },
                                    `Actualizó descuento a ${val}%`
                                  ).then((ok) => {
                                    if (ok && prevAbonado > 0) {
                                      const nuevoTot = round2(pedido.subtotalNeto * (1 - val / 100) * 1.21);
                                      if (nuevoTot < prevAbonado) {
                                        onNotify(
                                          `El monto cobrado se ajustó al nuevo total de ${formatARS(nuevoTot)}.`,
                                          'amber'
                                        );
                                      }
                                    }
                                  });
                                }
                              }}
                              className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums disabled:bg-zinc-50 dark:disabled:bg-[#141417] disabled:text-zinc-400 dark:disabled:text-zinc-500"
                            />
                          </div>
                        </div>
                      </div>
                    )
                  ) : (
                    <p className="text-[11px] text-zinc-400 dark:text-zinc-500 flex items-center gap-1.5">
                      <Lock className="w-3 h-3" strokeWidth={1.5} />
                      Los pagos y cheques solo pueden modificarse con perfil Administración.
                    </p>
                  )}
                </div>
              )}

              <div className="pt-5 border-t border-zinc-200/70 dark:border-[#26262B] space-y-3">
                <div className="flex items-center justify-between gap-2 flex-wrap">
                  <div className="flex items-center gap-2">
                    <h3 className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 uppercase tracking-wider">
                      Armado y envío
                    </h3>
                    {renderSaveBadge('logistica')}
                  </div>

                  {msgDespacho && (
                    <div className="inline-flex items-center gap-1">
                      {waUrlDespacho ? (
                        <>
                          <a
                            href={waUrlDespacho}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors"
                          >
                            <Send className="w-3 h-3" strokeWidth={1.5} />
                            Enviar aviso por WhatsApp
                          </a>
                          <button
                            type="button"
                            title="Copiar mensaje de envío"
                            aria-label="Copiar mensaje de envío"
                            onClick={() =>
                              onCopyText(`envio-${pedido.id}`, msgDespacho, 'Aviso de envío copiado.')
                            }
                            className="p-1.5 rounded-full bg-zinc-100 hover:bg-zinc-200/80 text-zinc-700 dark:bg-[#1B1B1F] dark:hover:bg-[#24242B] dark:text-zinc-200"
                          >
                            {copiadoId === `envio-${pedido.id}` ? (
                              <Check className="w-3 h-3 text-emerald-600 dark:text-emerald-400" strokeWidth={1.75} />
                            ) : (
                              <Copy className="w-3 h-3" strokeWidth={1.5} />
                            )}
                          </button>
                        </>
                      ) : (
                        <button
                          type="button"
                          onClick={() =>
                            onCopyText(
                              `envio-${pedido.id}`,
                              msgDespacho,
                              'Aviso de envío copiado para WhatsApp'
                            )
                          }
                          className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-medium bg-zinc-100 hover:bg-zinc-200/80 text-zinc-700 dark:bg-[#1B1B1F] dark:hover:bg-[#24242B] dark:text-zinc-200 transition-colors"
                        >
                          {copiadoId === `envio-${pedido.id}` ? (
                            <Check className="w-3 h-3 text-emerald-600 dark:text-emerald-400" strokeWidth={1.75} />
                          ) : (
                            <Send className="w-3 h-3" strokeWidth={1.5} />
                          )}
                          {copiadoId === `envio-${pedido.id}` ? 'Aviso copiado' : 'Copiar aviso de envío'}
                        </button>
                      )}
                    </div>
                  )}
                </div>

                <div className="space-y-3 text-xs">
                  <div>
                    <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1.5">
                      Estado de armado
                    </label>
                    {pedido.faltantesVariante > 0 ? (
                      <div className="px-3 py-2 rounded-xl bg-red-50 text-red-700 border border-red-200/70 dark:bg-red-500/10 dark:text-red-300 dark:border-red-500/25 font-medium">
                        Falta talle o color — definilo arriba para poder armar
                      </div>
                    ) : (
                      <div className="space-y-2">
                        <div className="grid grid-cols-3 p-1 rounded-xl bg-zinc-100 dark:bg-[#1B1B1F] border border-zinc-200/70 dark:border-[#26262B] gap-1">
                          {(['Por armar', 'En armado', 'Listo'] as const).map((est) => {
                            const activo = pedido.estadoArmado === est;
                            return (
                              <button
                                key={est}
                                type="button"
                                onClick={() => handleCambioArmado(est)}
                                className={`min-h-[38px] px-2 rounded-lg text-xs font-medium transition-colors ${
                                  activo
                                    ? 'bg-white text-[#111111] dark:bg-[#2A2A32] dark:text-white shadow-sm'
                                    : 'text-zinc-600 hover:text-[#111111] dark:text-zinc-400 dark:hover:text-[#F4F4F5]'
                                }`}
                              >
                                {est === 'Listo' ? 'Listo' : est}
                              </button>
                            );
                          })}
                        </div>

                        {confirmArmadoDestino && (
                          <div className="flex items-center justify-between gap-2 bg-amber-50 border border-amber-200 dark:bg-amber-500/10 dark:border-amber-500/25 rounded-xl px-3 py-2 text-[11px]">
                            <span className="text-amber-900 dark:text-amber-200">
                              El pedido figura {pedido.estadoEnvio}. ¿Cambiar armado a {confirmArmadoDestino}?
                            </span>
                            <div className="flex gap-1.5 shrink-0">
                              <button
                                type="button"
                                onClick={() => setConfirmArmadoDestino(null)}
                                className="px-2.5 py-0.5 rounded-full bg-white border border-zinc-200 text-zinc-700 dark:bg-[#1B1B1F] dark:border-[#2E2E35] dark:text-zinc-200"
                              >
                                No
                              </button>
                              <button
                                type="button"
                                onClick={() => {
                                  const d = confirmArmadoDestino;
                                  setConfirmArmadoDestino(null);
                                  void patchConEstado('logistica', { estadoArmado: d }, `Cambió armado a ${d}`);
                                }}
                                className="px-2.5 py-0.5 rounded-full bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-950 font-medium"
                              >
                                Sí
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  <div>
                    <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1.5">
                      Estado de envío
                    </label>
                    <div className="grid grid-cols-3 p-1 rounded-xl bg-zinc-100 dark:bg-[#1B1B1F] border border-zinc-200/70 dark:border-[#26262B] gap-1">
                      {(['Pendiente', 'Despachado', 'Entregado'] as const).map((est) => {
                        const activo = pedido.estadoEnvio === est;
                        const bloqueado = pedido.faltantesVariante > 0 && est !== 'Pendiente';
                        return (
                          <button
                            key={est}
                            type="button"
                            disabled={bloqueado}
                            title={bloqueado ? 'Definí los talles para poder despachar' : undefined}
                            onClick={() => handleCambioEnvio(est)}
                            className={`min-h-[38px] px-2 rounded-lg text-xs font-medium transition-colors disabled:opacity-35 disabled:cursor-not-allowed ${
                              activo
                                ? 'bg-white text-[#111111] dark:bg-[#2A2A32] dark:text-white shadow-sm'
                                : 'text-zinc-600 hover:text-[#111111] dark:text-zinc-400 dark:hover:text-[#F4F4F5]'
                            }`}
                          >
                            {est}
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Transporte</label>
                      <input
                        ref={transporteInputRef}
                        type="text"
                        list={datalistId}
                        value={draft.transporte}
                        onChange={(e) => setDraft((prev) => ({ ...prev, transporte: e.target.value }))}
                        onBlur={() => {
                          const val = draft.transporte.trim();
                          if (val !== pedido.transporte) {
                            void patchConEstado('logistica', { transporte: val }, `Transporte: ${val || 'vacío'}`);
                          }
                        }}
                        placeholder="Ej. Andreani, Cruz del Sur"
                        className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                      />
                    </div>

                    <div>
                      <div className="flex items-center justify-between mb-1">
                        <label className="text-[11px] text-zinc-500 dark:text-zinc-400">
                          N.º de seguimiento / guía
                        </label>
                        {draft.nroSeguimiento.trim() && (
                          <button
                            type="button"
                            onClick={() =>
                              onCopyText(
                                `guia-${pedido.id}`,
                                draft.nroSeguimiento.trim(),
                                'Guía de seguimiento copiada.'
                              )
                            }
                            className="inline-flex items-center gap-1 text-[10px] text-zinc-500 hover:text-[#111111] dark:text-zinc-400 dark:hover:text-[#F4F4F5]"
                          >
                            {copiadoId === `guia-${pedido.id}` ? (
                              <Check className="w-3 h-3 text-emerald-600 dark:text-emerald-400" strokeWidth={2} />
                            ) : (
                              <Copy className="w-3 h-3" strokeWidth={1.5} />
                            )}
                            {copiadoId === `guia-${pedido.id}` ? 'Copiada' : 'Copiar'}
                          </button>
                        )}
                      </div>
                      <input
                        type="text"
                        value={draft.nroSeguimiento}
                        onChange={(e) => setDraft((prev) => ({ ...prev, nroSeguimiento: e.target.value }))}
                        onBlur={() => {
                          const val = draft.nroSeguimiento.trim();
                          if (val !== pedido.nroSeguimiento) {
                            void patchConEstado(
                              'logistica',
                              { nroSeguimiento: val },
                              `Guía de seguimiento: ${val || 'vacía'}`
                            );
                          }
                        }}
                        placeholder="Código de guía"
                        className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Observaciones</label>
                    <textarea
                      rows={2}
                      value={draft.observaciones}
                      onChange={(e) => setDraft((prev) => ({ ...prev, observaciones: e.target.value }))}
                      onBlur={() => {
                        const val = draft.observaciones.trim();
                        if (val !== pedido.observaciones) {
                          void patchConEstado('logistica', { observaciones: val }, 'Actualizó observaciones');
                        }
                      }}
                      className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl"
                    />
                  </div>
                </div>

                {esDeposito && (
                  <div className="pt-2">
                    {!mostrarIncidenciaForm ? (
                      <button
                        type="button"
                        onClick={() => setMostrarIncidenciaForm(true)}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-amber-50 border border-amber-200 text-amber-800 hover:bg-amber-100/80 dark:bg-amber-500/[0.12] dark:border-amber-500/25 dark:text-amber-300 dark:hover:bg-amber-500/20"
                      >
                        <AlertTriangle className="w-3.5 h-3.5" strokeWidth={1.5} />
                        Reportar faltante / problema
                      </button>
                    ) : (
                      <div className="bg-amber-50/70 border border-amber-200 dark:bg-amber-500/10 dark:border-amber-500/25 rounded-xl p-3 space-y-2 text-xs">
                        <label className="block text-[11px] font-medium text-amber-900 dark:text-amber-200">
                          Detalle para Administración
                        </label>
                        <div className="flex flex-wrap gap-1.5">
                          {QUICK_INCIDENCIA_CHIPS.map((chip) => (
                            <button
                              key={chip}
                              type="button"
                              onClick={() =>
                                setIncidenciaTexto((prev) => (prev.trim() ? `${prev.trim()} · ${chip}` : chip))
                              }
                              className="px-2.5 py-1 rounded-full text-[11px] bg-white border border-amber-200 text-amber-900 hover:bg-amber-100 dark:bg-[#1B1B1F] dark:border-amber-500/30 dark:text-amber-200"
                            >
                              + {chip}
                            </button>
                          ))}
                        </div>
                        <textarea
                          rows={2}
                          value={incidenciaTexto}
                          onChange={(e) => setIncidenciaTexto(e.target.value)}
                          placeholder="Ej. Faltan 2 Chalecos Shark T1 en estantería..."
                          className="w-full px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-amber-200 dark:border-amber-500/30 rounded-lg placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                        />
                        <div className="flex justify-end gap-2">
                          <button
                            type="button"
                            onClick={() => setMostrarIncidenciaForm(false)}
                            className="px-3 py-1.5 rounded-full bg-white border border-zinc-200 text-zinc-600 dark:bg-[#1B1B1F] dark:border-[#2E2E35] dark:text-zinc-300"
                          >
                            Cancelar
                          </button>
                          <button
                            type="button"
                            disabled={enviandoIncidencia}
                            onClick={() => void handleEnviarIncidencia()}
                            className="px-3.5 py-1.5 rounded-full bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-950 font-medium disabled:opacity-40"
                          >
                            {enviandoIncidencia ? 'Enviando...' : 'Enviar aviso'}
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {(pedido.lineas.length > 0 || esAdmin) && (
                <div className="pt-5 border-t border-zinc-200/70 dark:border-[#26262B] space-y-3">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <h3 className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 uppercase tracking-wider">
                      Detalle de artículos ({editandoArticulos ? lineasEdicion.length : pedido.lineas.length} ·{' '}
                      {editandoArticulos ? (previewAjuste?.unidadesNuevas ?? 0) : totalUnidades} u.)
                    </h3>

                    {!esDeposito &&
                      !editandoArticulos &&
                      (puedeAjustarArticulos ? (
                        <button
                          type="button"
                          onClick={handleEntrarAjusteArticulos}
                          className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-medium bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 text-[#111111] dark:bg-[#1B1B1F] dark:hover:bg-[#24242B] dark:border-[#2E2E35] dark:text-[#F4F4F5] transition-colors"
                        >
                          <Edit3 className="w-3 h-3" strokeWidth={1.5} />
                          {pedido.lineas.length === 0 ? 'Cargar artículos' : 'Ajustar cantidades'}
                        </button>
                      ) : (
                        <span className="text-[11px] text-zinc-400 dark:text-zinc-500 flex items-center gap-1">
                          <Lock className="w-3 h-3 shrink-0" strokeWidth={1.5} />
                          Ya fue enviado a Depósito o tiene cobros registrados; pedile el ajuste a Administración.
                        </span>
                      ))}
                  </div>

                  {ultimoAjusteGuardado && !esDeposito && !editandoArticulos && (
                    <div className="bg-emerald-50/80 border border-emerald-200 dark:bg-emerald-500/10 dark:border-emerald-500/25 rounded-2xl p-3.5 space-y-2.5 text-xs">
                      <div className="flex items-start justify-between gap-2">
                        <div className="space-y-0.5">
                          <div className="font-semibold text-emerald-900 dark:text-emerald-200 flex items-center gap-1.5">
                            <Check className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400" strokeWidth={2} />
                            Ajuste guardado · {ultimoAjusteGuardado.unidadesAnteriores} → {ultimoAjusteGuardado.unidadesNuevas} u.
                          </div>
                          <div className="text-[11px] text-emerald-800/90 dark:text-emerald-300/90 tabular-nums">
                            Total: {formatARS(ultimoAjusteGuardado.totalAnterior)} →{' '}
                            <strong>{formatARS(ultimoAjusteGuardado.totalNuevo)}</strong>
                            {ultimoAjusteGuardado.saldoAFavorGenerado > 0.005
                              ? ` · Saldo a favor: ${formatARS(ultimoAjusteGuardado.saldoAFavorGenerado)}`
                              : pedido.saldoPendiente > 0.005
                                ? ` · Saldo pendiente: ${formatARS(pedido.saldoPendiente)}`
                                : ''}
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={() => setUltimoAjusteGuardado(null)}
                          aria-label="Cerrar aviso"
                          className="p-1 text-emerald-700/70 hover:text-emerald-900 dark:text-emerald-300/70 dark:hover:text-emerald-100"
                        >
                          <X className="w-3.5 h-3.5" strokeWidth={1.75} />
                        </button>
                      </div>

                      <div className="flex flex-wrap items-center gap-1.5">
                        {waUrlAjusteWhatsapp ? (
                          <>
                            <a
                              href={waUrlAjusteWhatsapp}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[11px] font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors"
                            >
                              <MessageCircle className="w-3.5 h-3.5" strokeWidth={1.75} />
                              Avisar por WhatsApp
                            </a>
                            <button
                              type="button"
                              title="Copiar mensaje"
                              aria-label="Copiar aviso de ajuste para WhatsApp"
                              onClick={() =>
                                onCopyText(
                                  `ajuste-wsp-${pedido.id}`,
                                  msgAjusteWhatsapp,
                                  'Aviso de ajuste copiado para WhatsApp.'
                                )
                              }
                              className="p-1.5 rounded-full bg-white border border-emerald-200 text-emerald-800 hover:bg-emerald-50 dark:bg-[#1B1B1F] dark:border-emerald-500/30 dark:text-emerald-200"
                            >
                              {copiadoId === `ajuste-wsp-${pedido.id}` ? (
                                <Check className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400" strokeWidth={1.75} />
                              ) : (
                                <Copy className="w-3.5 h-3.5" strokeWidth={1.5} />
                              )}
                            </button>
                          </>
                        ) : (
                          <button
                            type="button"
                            onClick={() =>
                              onCopyText(
                                `ajuste-wsp-${pedido.id}`,
                                msgAjusteWhatsapp,
                                'Aviso de ajuste copiado para WhatsApp.'
                              )
                            }
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[11px] font-medium bg-white border border-emerald-200 text-emerald-800 hover:bg-emerald-50 dark:bg-[#1B1B1F] dark:border-emerald-500/30 dark:text-emerald-200"
                          >
                            {copiadoId === `ajuste-wsp-${pedido.id}` ? (
                              <Check className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400" strokeWidth={1.75} />
                            ) : (
                              <MessageCircle className="w-3.5 h-3.5" strokeWidth={1.5} />
                            )}
                            {copiadoId === `ajuste-wsp-${pedido.id}` ? 'Aviso copiado' : 'Copiar aviso para WhatsApp'}
                          </button>
                        )}
                      </div>
                    </div>
                  )}

                  {!editandoArticulos ? (
                    pedido.lineas.length > 0 ? (
                      <div className="divide-y divide-zinc-200/60 dark:divide-[#26262B]/80 border border-zinc-200/80 dark:border-[#26262B] rounded-xl">
                        {pedido.lineas.map((l) => {
                          const prod = catalogoCompleto.find((p) => p.idGrupo === l.idGrupo);
                          const variantes = prod ? prod.variantes : l.desglose;
                          const puedeEditarTalles = !esDeposito && !l.requiereRevision && variantes.length >= 2;
                          return (
                            <div key={l.idLinea} className="px-3.5 py-2.5 text-xs space-y-2">
                              <div className="flex justify-between gap-2 items-start">
                                <div>
                                  <div className="font-medium text-[#111111] dark:text-[#F4F4F5]">{l.nombreProducto}</div>
                                  <div className="text-[11px] text-zinc-400 dark:text-zinc-500 tabular-nums">
                                    {l.requiereRevision
                                      ? `Sin definir (${l.cantidadTotal} u.)`
                                      : l.desglose
                                          .filter((d) => d.cantidad > 0)
                                          .map((d) => `${d.codigo} · ${d.etiqueta}: ${d.cantidad}`)
                                          .join(' | ')}
                                  </div>
                                </div>
                                <div className="flex items-center gap-2.5 shrink-0">
                                  {puedeEditarTalles && editandoLineaId !== l.idLinea && (
                                    <button
                                      type="button"
                                      onClick={() => setEditandoLineaId(l.idLinea)}
                                      className="text-[11px] text-zinc-400 hover:text-[#111111] dark:text-zinc-500 dark:hover:text-[#F4F4F5] inline-flex items-center gap-1"
                                    >
                                      <Edit3 className="w-3 h-3" strokeWidth={1.5} />
                                      Editar talles
                                    </button>
                                  )}
                                  <div className="font-semibold text-[#111111] dark:text-[#F4F4F5] tabular-nums whitespace-nowrap">
                                    {esDeposito
                                      ? `${l.cantidadTotal} u.`
                                      : formatARS(round2(l.precioUnitarioNeto * l.cantidadTotal))}
                                  </div>
                                </div>
                              </div>

                              {editandoLineaId === l.idLinea && renderVariantEditor(l, 'edicion', false)}
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <p className="text-[11px] text-zinc-400 dark:text-zinc-500">
                        Pedido histórico importado sin desglose de artículos ({pedido.detalleManualImportado || '—'}).
                      </p>
                    )
                  ) : (
                    <div className="space-y-3">
                      <div className="divide-y divide-zinc-200/70 dark:divide-[#26262B] border border-zinc-200 dark:border-[#2E2E35] rounded-2xl bg-[#FAFAFA] dark:bg-[#1B1B1F]">
                        {lineasEdicion.length === 0 ? (
                          <div className="p-4 text-center text-xs text-red-600 dark:text-red-400">
                            Un pedido debe tener al menos 1 artículo; para anularlo completo usá Eliminar pedido.
                          </div>
                        ) : (
                          lineasEdicion.map((l) => {
                            const prod = catalogoCompleto.find((p) => p.idGrupo === l.idGrupo);
                            const variantesBase = prod && prod.variantes.length > 0 ? prod.variantes : l.desglose;
                            const esMultiResuelto = !l.requiereRevision && l.estado === 'OK' && variantesBase.length >= 2;
                            const esConfirmQuitar = confirmQuitarIdGrupo === l.idGrupo;

                            return (
                              <div key={l.idLinea} className="p-3.5 space-y-2.5 text-xs">
                                <div className="flex items-start justify-between gap-2">
                                  <div className="min-w-0">
                                    <div className="font-semibold text-[#111111] dark:text-[#F4F4F5] truncate">
                                      {l.nombreProducto}
                                    </div>
                                    <div className="text-[11px] text-zinc-500 dark:text-zinc-400 tabular-nums">
                                      {formatARSCompacto(l.precioUnitarioNeto)} c/u neto · Subtotal:{' '}
                                      <strong className="text-[#111111] dark:text-[#F4F4F5]">
                                        {formatARS(round2(l.precioUnitarioNeto * l.cantidadTotal))}
                                      </strong>
                                    </div>
                                  </div>

                                  <div className="flex items-center gap-1.5 shrink-0">
                                    {!esMultiResuelto && (
                                      <div className="inline-flex items-center rounded-xl border border-zinc-200 dark:border-[#2E2E35] bg-white dark:bg-[#141417]">
                                        <button
                                          type="button"
                                          onClick={() =>
                                            handleAjustarCantidadSimpleEdicion(l.idGrupo, l.cantidadTotal - 1)
                                          }
                                          aria-label={`Restar ${l.nombreProducto}`}
                                          className="p-1.5 text-zinc-600 hover:text-[#111111] dark:text-zinc-300 dark:hover:text-white"
                                        >
                                          <Minus className="w-3 h-3" strokeWidth={1.75} />
                                        </button>
                                        <input
                                          type="number"
                                          min={1}
                                          value={l.cantidadTotal}
                                          onWheel={(e) => e.currentTarget.blur()}
                                          onChange={(e) => {
                                            const n = parseInt(e.target.value, 10);
                                            if (Number.isFinite(n) && n >= 1) {
                                              handleAjustarCantidadSimpleEdicion(l.idGrupo, n);
                                            }
                                          }}
                                          className="w-11 text-center text-xs font-semibold bg-transparent text-[#111111] dark:text-[#F4F4F5] tabular-nums focus:outline-none"
                                        />
                                        <button
                                          type="button"
                                          onClick={() =>
                                            handleAjustarCantidadSimpleEdicion(l.idGrupo, l.cantidadTotal + 1)
                                          }
                                          aria-label={`Sumar ${l.nombreProducto}`}
                                          className="p-1.5 text-zinc-600 hover:text-[#111111] dark:text-zinc-300 dark:hover:text-white"
                                        >
                                          <Plus className="w-3 h-3" strokeWidth={1.75} />
                                        </button>
                                      </div>
                                    )}

                                    <button
                                      type="button"
                                      onClick={() => handleQuitarLineaEdicion(l.idGrupo)}
                                      title="Quitar artículo"
                                      aria-label={`Quitar ${l.nombreProducto}`}
                                      className="p-1.5 rounded-lg text-zinc-400 hover:text-red-600 hover:bg-red-50 dark:text-zinc-500 dark:hover:text-red-400 dark:hover:bg-red-500/10 transition-colors"
                                    >
                                      <Trash2 className="w-3.5 h-3.5" strokeWidth={1.5} />
                                    </button>
                                  </div>
                                </div>

                                {l.requiereRevision && (
                                  <div className="text-[11px] text-amber-700 dark:text-amber-300">
                                    Falta definir talle o color ({l.cantidadTotal} u. totales)
                                  </div>
                                )}

                                {esMultiResuelto && (
                                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                                    {variantesBase.map((v) => {
                                      const qtyVar =
                                        l.desglose.find((d) => d.codigo === v.codigo)?.cantidad ?? 0;
                                      return (
                                        <div
                                          key={v.codigo}
                                          className="flex items-center justify-between gap-2 rounded-xl bg-white dark:bg-[#141417] border border-zinc-200/80 dark:border-[#26262B] px-2.5 py-1.5"
                                        >
                                          <span className="text-[11px] text-zinc-600 dark:text-zinc-300 truncate">
                                            <strong className="text-[#111111] dark:text-[#F4F4F5]">{v.codigo}</strong> ·{' '}
                                            {v.etiqueta}
                                          </span>
                                          <div className="inline-flex items-center rounded-lg border border-zinc-200 dark:border-[#2E2E35] bg-[#FAFAFA] dark:bg-[#1B1B1F] shrink-0">
                                            <button
                                              type="button"
                                              onClick={() =>
                                                handleAjustarVarianteEdicion(l.idGrupo, v.codigo, qtyVar - 1)
                                              }
                                              aria-label={`Restar ${v.etiqueta}`}
                                              className="p-1 text-zinc-600 hover:text-[#111111] dark:text-zinc-300 dark:hover:text-white"
                                            >
                                              <Minus className="w-3 h-3" strokeWidth={1.75} />
                                            </button>
                                            <input
                                              type="number"
                                              min={0}
                                              value={qtyVar}
                                              onWheel={(e) => e.currentTarget.blur()}
                                              onChange={(e) => {
                                                const n = parseInt(e.target.value, 10);
                                                handleAjustarVarianteEdicion(
                                                  l.idGrupo,
                                                  v.codigo,
                                                  Number.isFinite(n) ? Math.max(0, n) : 0
                                                );
                                              }}
                                              className="w-9 text-center text-xs font-semibold bg-transparent text-[#111111] dark:text-[#F4F4F5] tabular-nums focus:outline-none"
                                            />
                                            <button
                                              type="button"
                                              onClick={() =>
                                                handleAjustarVarianteEdicion(l.idGrupo, v.codigo, qtyVar + 1)
                                              }
                                              aria-label={`Sumar ${v.etiqueta}`}
                                              className="p-1 text-zinc-600 hover:text-[#111111] dark:text-zinc-300 dark:hover:text-white"
                                            >
                                              <Plus className="w-3 h-3" strokeWidth={1.75} />
                                            </button>
                                          </div>
                                        </div>
                                      );
                                    })}
                                  </div>
                                )}

                                {esConfirmQuitar && (
                                  <div className="flex items-center justify-between gap-2 rounded-xl bg-red-50/90 border border-red-200 dark:bg-red-500/10 dark:border-red-500/30 px-3 py-2 text-[11px]">
                                    <span className="text-red-800 dark:text-red-200 font-medium">
                                      Es la última unidad. ¿Quitar artículo del pedido?
                                    </span>
                                    <div className="flex items-center gap-1.5 shrink-0">
                                      <button
                                        type="button"
                                        onClick={() => setConfirmQuitarIdGrupo(null)}
                                        className="px-2.5 py-0.5 rounded-full bg-white border border-zinc-200 text-zinc-700 dark:bg-[#141417] dark:border-[#2E2E35] dark:text-zinc-200"
                                      >
                                        No
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => handleQuitarLineaEdicion(l.idGrupo)}
                                        className="px-2.5 py-0.5 rounded-full bg-red-600 text-white font-semibold"
                                      >
                                        Quitar artículo
                                      </button>
                                    </div>
                                  </div>
                                )}
                              </div>
                            );
                          })
                        )}
                      </div>

                      <div>
                        <button
                          type="button"
                          onClick={() => {
                            setMostrarBuscadorReemplazo((prev) => !prev);
                            setReemplazoPendienteProd(null);
                            setMostrarAdhocDrawer(false);
                          }}
                          className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-full text-xs font-medium bg-white hover:bg-zinc-50 border border-zinc-200 text-[#111111] dark:bg-[#1B1B1F] dark:hover:bg-[#24242B] dark:border-[#2E2E35] dark:text-[#F4F4F5] transition-colors"
                        >
                          <Plus className="w-3.5 h-3.5" strokeWidth={1.75} />
                          {mostrarBuscadorReemplazo ? 'Cerrar reemplazos' : '+ Agregar reemplazo'}
                        </button>
                      </div>

                      {mostrarBuscadorReemplazo && (
                        <div className="rounded-2xl border border-zinc-200 dark:border-[#2E2E35] bg-white dark:bg-[#141417] p-3.5 space-y-2.5 text-xs">
                          <div className="flex items-center justify-between gap-2">
                            <div className="relative flex-1">
                              <Search
                                className="w-3.5 h-3.5 text-zinc-400 dark:text-zinc-500 absolute left-3 top-1/2 -translate-y-1/2"
                                strokeWidth={1.5}
                              />
                              <input
                                type="text"
                                value={busquedaReemplazo}
                                onChange={(e) => {
                                  setBusquedaReemplazo(e.target.value);
                                  setReemplazoPendienteProd(null);
                                }}
                                placeholder="Buscar en catálogo o Pilusos por nombre, código o rubro..."
                                className="w-full pl-8 pr-3 py-1.5 text-xs bg-[#FAFAFA] dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
                              />
                            </div>
                            <button
                              type="button"
                              onClick={() => {
                                setMostrarAdhocDrawer((prev) => !prev);
                                setReemplazoPendienteProd(null);
                              }}
                              className="px-2.5 py-1.5 rounded-xl text-[11px] font-medium border border-zinc-200 dark:border-[#2E2E35] text-zinc-600 dark:text-zinc-300 hover:bg-zinc-50 dark:hover:bg-[#24242B] shrink-0"
                            >
                              Fuera de planilla
                            </button>
                          </div>

                          {mostrarAdhocDrawer ? (
                            <div className="rounded-xl bg-[#FAFAFA] dark:bg-[#1B1B1F] border border-zinc-200/80 dark:border-[#26262B] p-3 space-y-2">
                              <div className="font-medium text-[#111111] dark:text-[#F4F4F5]">
                                Artículo fuera de planilla (solo este pedido)
                              </div>
                              <div className="grid grid-cols-2 gap-2">
                                <input
                                  type="text"
                                  value={adhocNombre}
                                  onChange={(e) => setAdhocNombre(e.target.value)}
                                  placeholder="Nombre del producto"
                                  className="col-span-2 px-2.5 py-1.5 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg"
                                />
                                <input
                                  type="text"
                                  value={adhocRubro}
                                  onChange={(e) => setAdhocRubro(e.target.value)}
                                  placeholder="Rubro"
                                  className="px-2.5 py-1.5 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg uppercase"
                                />
                                <input
                                  type="number"
                                  min={0}
                                  step="0.01"
                                  value={adhocPrecio}
                                  onWheel={(e) => e.currentTarget.blur()}
                                  onChange={(e) => setAdhocPrecio(e.target.value)}
                                  placeholder="Precio neto ($)"
                                  className="px-2.5 py-1.5 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg tabular-nums"
                                />
                              </div>
                              <div className="flex items-center justify-between gap-2">
                                <div className="flex items-center gap-1.5">
                                  <span className="text-[11px] text-zinc-500 dark:text-zinc-400">Cant.:</span>
                                  <input
                                    type="number"
                                    min={1}
                                    value={adhocCantidad}
                                    onWheel={(e) => e.currentTarget.blur()}
                                    onChange={(e) => setAdhocCantidad(e.target.value)}
                                    className="w-16 px-2 py-1 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg tabular-nums"
                                  />
                                </div>
                                <button
                                  type="button"
                                  onClick={handleAgregarAdhocEnDrawer}
                                  className="px-3 py-1 rounded-full bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-950 font-medium"
                                >
                                  Sumar artículo
                                </button>
                              </div>
                            </div>
                          ) : reemplazoPendienteProd ? (
                            <div className="rounded-xl bg-[#FAFAFA] dark:bg-[#1B1B1F] border border-zinc-200/80 dark:border-[#26262B] p-3 space-y-2.5">
                              <div className="flex items-center justify-between gap-2">
                                <span className="font-semibold text-[#111111] dark:text-[#F4F4F5]">
                                  {reemplazoPendienteProd.nombre}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => setReemplazoPendienteProd(null)}
                                  className="text-[11px] text-zinc-500 hover:text-zinc-900 dark:text-zinc-400 dark:hover:text-white"
                                >
                                  Cambiar producto
                                </button>
                              </div>
                              <div className="grid grid-cols-2 gap-2">
                                {reemplazoPendienteProd.variantes.map((v) => (
                                  <div key={v.codigo}>
                                    <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                                      {v.codigo} · {v.etiqueta}
                                    </label>
                                    <input
                                      type="number"
                                      min={0}
                                      value={reemplazoVarsQty[v.codigo] ?? '0'}
                                      onWheel={(e) => e.currentTarget.blur()}
                                      onChange={(e) =>
                                        setReemplazoVarsQty((prev) => ({
                                          ...prev,
                                          [v.codigo]: e.target.value
                                        }))
                                      }
                                      className="w-full px-2.5 py-1.5 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg tabular-nums"
                                    />
                                  </div>
                                ))}
                              </div>
                              <div className="flex flex-wrap items-center justify-end gap-2">
                                {pedido.estadoEnvio === 'Pendiente' && (
                                  <button
                                    type="button"
                                    onClick={() => handleConfirmarReemplazoVariantes(true)}
                                    className="px-3 py-1.5 rounded-full border border-zinc-200 dark:border-[#2E2E35] text-zinc-600 dark:text-zinc-300"
                                  >
                                    Agregar pendiente de talle
                                  </button>
                                )}
                                <button
                                  type="button"
                                  onClick={() => handleConfirmarReemplazoVariantes(false)}
                                  className="px-3.5 py-1.5 rounded-full bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-950 font-medium"
                                >
                                  Sumar al pedido
                                </button>
                              </div>
                            </div>
                          ) : (
                            <div className="max-h-44 overflow-y-auto divide-y divide-zinc-100 dark:divide-[#26262B] border border-zinc-200/70 dark:border-[#26262B] rounded-xl">
                              {resultadosBusquedaReemplazo.length === 0 ? (
                                <div className="px-3 py-2.5 text-[11px] text-zinc-400 dark:text-zinc-500">
                                  No hay productos con ese criterio.
                                </div>
                              ) : (
                                resultadosBusquedaReemplazo.map((prod) => (
                                  <button
                                    key={prod.idGrupo}
                                    type="button"
                                    onClick={() => handleSeleccionarReemplazoCatalogo(prod)}
                                    className="w-full px-3 py-2 flex items-center justify-between gap-2 text-left hover:bg-zinc-50 dark:hover:bg-[#1B1B1F] transition-colors"
                                  >
                                    <div className="min-w-0">
                                      <div className="font-medium text-[#111111] dark:text-[#F4F4F5] truncate flex items-center gap-1.5">
                                        <span className="truncate">{prod.nombre}</span>
                                        {prod.esCustom && (
                                          <span className="px-1.5 py-0.5 rounded bg-[#FCE7F3] text-[#111111] dark:bg-[#F4B6CE]/15 dark:text-[#FCE7F3] text-[10px]">
                                            Catálogo extra
                                          </span>
                                        )}
                                      </div>
                                      <div className="text-[10px] text-zinc-400 dark:text-zinc-500 tabular-nums">
                                        {prod.variantes.map((v) => v.codigo).join(' / ')} · {prod.rubro}
                                      </div>
                                    </div>
                                    <span className="text-[11px] font-medium text-zinc-700 dark:text-zinc-300 shrink-0 tabular-nums">
                                      {formatARSCompacto(prod.precioNeto)} neto
                                    </span>
                                  </button>
                                ))
                              )}
                            </div>
                          )}
                        </div>
                      )}

                      {previewAjuste && (
                        <div className="rounded-2xl border border-zinc-200 dark:border-[#2E2E35] bg-white dark:bg-[#141417] p-4 space-y-3 text-xs">
                          <div className="flex flex-wrap items-center justify-between gap-2 tabular-nums">
                            <div className="text-zinc-600 dark:text-zinc-300">
                              Unidades: <strong className="text-[#111111] dark:text-[#F4F4F5]">{totalUnidades} u.</strong> →{' '}
                              <strong className="text-[#111111] dark:text-[#F4F4F5]">{previewAjuste.unidadesNuevas} u.</strong>{' '}
                              {previewAjuste.deltaUnidades !== 0 && (
                                <span
                                  className={
                                    previewAjuste.deltaUnidades > 0
                                      ? 'text-emerald-700 dark:text-emerald-400'
                                      : 'text-amber-700 dark:text-amber-400'
                                  }
                                >
                                  ({previewAjuste.deltaUnidades > 0 ? '+' : ''}
                                  {previewAjuste.deltaUnidades})
                                </span>
                              )}
                            </div>
                            <div className="tabular-nums">
                              Total Facturado:{' '}
                              <span className="text-zinc-400 dark:text-zinc-500">
                                {formatARS(pedido.totalFacturado)}
                              </span>{' '}
                              →{' '}
                              <strong
                                className={
                                  previewAjuste.deltaTotalCents < 0
                                    ? 'text-amber-700 dark:text-amber-400'
                                    : previewAjuste.deltaTotalCents > 0
                                      ? 'text-emerald-700 dark:text-emerald-400'
                                      : 'text-[#111111] dark:text-[#F4F4F5]'
                                }
                              >
                                {formatARS(previewAjuste.previewPedido.totalFacturado)}
                              </strong>
                            </div>
                          </div>

                          {previewAjuste.saldoAFavor > 0.005 && (
                            <div className="rounded-xl bg-amber-50 border border-amber-300 dark:bg-amber-500/15 dark:border-amber-500/35 px-3 py-2 text-[11px] font-medium text-amber-900 dark:text-amber-200">
                              El nuevo total ({formatARS(previewAjuste.previewPedido.totalFacturado)}) es menor a lo ya cobrado (
                              {formatARS(pedido.montoAbonado)}). Queda un saldo a favor del cliente de{' '}
                              <strong>{formatARS(previewAjuste.saldoAFavor)}</strong>.
                            </div>
                          )}

                          {previewAjuste.diffCheque > 0.005 && (
                            <div className="rounded-xl bg-amber-50/90 border border-amber-200 dark:bg-amber-500/10 dark:border-amber-500/30 px-3 py-2 text-[11px] text-amber-900 dark:text-amber-200">
                              Revisar el cheque recibido por una diferencia de{' '}
                              <strong>{formatARS(previewAjuste.diffCheque)}</strong>.
                            </div>
                          )}

                          {pedido.medioPago === 'Cheque' &&
                            !pedido.chequeAcreditado &&
                            previewAjuste.diffCheque <= 0.005 &&
                            previewAjuste.previewPedido.deudaSinRespaldo > 0.005 && (
                              <div className="rounded-xl bg-amber-50/90 border border-amber-200 dark:bg-amber-500/10 dark:border-amber-500/30 px-3 py-2 text-[11px] text-amber-900 dark:text-amber-200">
                                Quedarán <strong>{formatARS(previewAjuste.previewPedido.deudaSinRespaldo)}</strong> sin cubrir por el cheque actual.
                              </div>
                            )}

                          {previewAjuste.pasoAPendiente && (
                            <div className="rounded-xl bg-zinc-50 border border-zinc-200 dark:bg-[#1B1B1F] dark:border-[#2E2E35] px-3 py-2 text-[11px] text-zinc-700 dark:text-zinc-300">
                              Quedará un saldo pendiente a cobrar de{' '}
                              <strong>{formatARS(previewAjuste.previewPedido.saldoPendiente)}</strong>.
                            </div>
                          )}

                          {asignacionesDeposito.length > 0 && (
                            <div className="text-[11px] text-amber-700 dark:text-amber-300">
                              Depósito verá este pedido como &apos;Actualizado&apos; para revisar la hoja.
                            </div>
                          )}

                          {pedido.incidenciaDeposito && esAdmin && (
                            <label className="flex items-center gap-2 text-[11px] text-zinc-700 dark:text-zinc-300 cursor-pointer select-none">
                              <input
                                type="checkbox"
                                checked={resolverIncidenciaCheck}
                                onChange={(e) => setResolverIncidenciaCheck(e.target.checked)}
                                className="w-4 h-4 rounded border-zinc-300 dark:border-zinc-600 accent-zinc-900 dark:accent-[#F4B6CE]"
                              />
                              <span>Marcar el aviso de Depósito como resuelto</span>
                            </label>
                          )}

                          {pedido.estadoEnvio !== 'Pendiente' && (
                            <label className="flex items-center gap-2 text-[11px] text-amber-800 dark:text-amber-300 font-medium cursor-pointer select-none">
                              <input
                                type="checkbox"
                                checked={confirmSobreDespachadoCheck}
                                onChange={(e) => setConfirmSobreDespachadoCheck(e.target.checked)}
                                className="w-4 h-4 rounded border-amber-400 dark:border-amber-500 accent-amber-600"
                              />
                              <span>Confirmar ajuste sobre pedido {pedido.estadoEnvio}</span>
                            </label>
                          )}

                          <div className="flex justify-end gap-2 pt-1">
                            <button
                              type="button"
                              onClick={handleCancelarAjusteArticulos}
                              className="px-3.5 py-2 rounded-full border border-zinc-200 dark:border-[#2E2E35] bg-white dark:bg-[#1B1B1F] text-zinc-600 dark:text-zinc-300"
                            >
                              Cancelar
                            </button>
                            <button
                              type="button"
                              disabled={
                                busyAjusteArticulos ||
                                !hayCambiosArticulos ||
                                lineasEdicion.length === 0 ||
                                (pedido.estadoEnvio !== 'Pendiente' && !confirmSobreDespachadoCheck)
                              }
                              onClick={() => void handleGuardarAjusteArticulos()}
                              className="px-4 py-2 rounded-full bg-zinc-900 hover:bg-zinc-800 text-white dark:bg-zinc-100 dark:hover:bg-white dark:text-zinc-950 font-medium transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
                            >
                              {busyAjusteArticulos ? 'Guardando...' : 'Guardar cambios'}
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}

              <div className="pt-5 border-t border-zinc-200/70 dark:border-[#26262B] space-y-2.5">
                <h3 className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 uppercase tracking-wider">
                  Historial de cambios
                </h3>
                <ul className="space-y-1.5 text-[11px] text-zinc-500 dark:text-zinc-400">
                  {pedido.historial.map((h, i) => (
                    <li key={i} className="flex items-start gap-2">
                      <Clock className="w-3 h-3 text-zinc-400 dark:text-zinc-500 mt-0.5 shrink-0" strokeWidth={1.5} />
                      <span>
                        <strong className="text-zinc-700 dark:text-zinc-200 font-medium">{h.fechaHora}</strong> · {h.usuario} ({h.rol}):{' '}
                        {h.accion}
                      </span>
                    </li>
                  ))}
                </ul>
              </div>
            </>
          )}
        </div>
      </aside>
    </div>
    {mostrarHojaImprimible && typeof document !== 'undefined'
      ? createPortal(
          <div className="hidden print:block mibi-print-portal">{renderHojaDeposito(false)}</div>,
          document.body
        )
      : null}
    </>
  );
}
