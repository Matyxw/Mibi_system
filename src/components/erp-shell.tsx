'use client';

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  BookOpen,
  Download,
  LogOut,
  Package,
  Plus,
  RotateCcw,
  Search,
  Undo2,
  Upload,
  Users,
  X
} from 'lucide-react';
import { logoutAction } from '@/app/actions/auth-actions';
import { listarProductosCustomAction } from '@/app/actions/catalogo-actions';
import {
  batchPedidosAction,
  crearPedidoAction,
  enviarPedidosACuentasAction,
  importarPlanillaAction,
  listarPedidosAction,
  marcarPedidoVistoAction,
  patchPedidoAction,
  reiniciarSistemaAction,
  reportarIncidenciaDepositoAction,
  resolverIncidenciaAction,
  restorePedidosAction,
  retirarAsignacionAction,
  softDeletePedidosAction
} from '@/app/actions/pedidos-actions';
import { listarCuentasAsignablesAction } from '@/app/actions/usuarios-actions';
import { CATALOGO_MIBI, MIBI_BRAND } from '@/lib/catalogo-mibi';
import type { EnviarCuentasResult, SyncStatusResult } from '@/lib/dal/pedidos';
import { aplicarResolucionLinea } from '@/lib/excel-parser';
import {
  addDaysISO,
  calcularResumenFinanciero,
  formatARS,
  formatARSCompacto,
  formatShortDate,
  getTodayISO,
  round2
} from '@/lib/finanzas';
import {
  isDepositoRole,
  type AdvertenciaExcel,
  type BatchPanelActivo,
  type FiltroPedidos,
  type LineaPedido,
  type MedioPago,
  type Pedido,
  type PrioridadAsignacion,
  type ProductoCatalogo,
  type ProductoCustomRecord,
  type ResumenFinanciero,
  type SessionUser,
  type ToastState,
  type ToneStatus,
  type UsuarioEquipo,
  type VistaActiva
} from '@/lib/types';
import { CatalogoModal } from './catalogo-modal';
import { EquipoModal } from './equipo-modal';
import { FloatingBatchBar } from './floating-batch-bar';
import { MibiLogo } from './mibi-logo';
import { crearFormInicial, NuevoPedidoView, type FormNuevoPedidoState } from './nuevo-pedido-view';
import { PedidoDrawer } from './pedido-drawer';
import { PedidosTable } from './pedidos-table';

const LS_FILTERS_KEY = 'mibi_erp_filtros_v3';

interface PendingUndoDelete {
  pedidos: Pedido[];
  timerId: number;
}

const copyTextSafe = (text: string): void => {
  if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
    void navigator.clipboard.writeText(text).catch(() => {});
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

const descargarResponseBlob = async (res: Response, fallbackName: string): Promise<void> => {
  const blob = await res.blob();
  const disp = res.headers.get('Content-Disposition') ?? '';
  const match = disp.match(/filename="([^"]+)"/);
  const filename = match?.[1] ?? fallbackName;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.setTimeout(() => URL.revokeObjectURL(url), 1500);
};

interface ErpShellProps {
  session: SessionUser;
  initialPedidos: Pedido[];
  initialCatalogoCustom: ProductoCustomRecord[];
  initialCuentasAsignables: UsuarioEquipo[];
  isLocalDev: boolean;
}

export function ErpShell({
  session,
  initialPedidos,
  initialCatalogoCustom,
  initialCuentasAsignables,
  isLocalDev
}: ErpShellProps): React.JSX.Element {
  const esAdmin = session.rol === 'ADMIN';
  const esDeposito = isDepositoRole(session.rol);

  const [pedidos, setPedidos] = useState<Pedido[]>(initialPedidos);
  const [productosCustom, setProductosCustom] = useState<ProductoCustomRecord[]>(initialCatalogoCustom);
  const [cuentasAsignables, setCuentasAsignables] = useState<UsuarioEquipo[]>(initialCuentasAsignables);
  const [vista, setVista] = useState<VistaActiva>('pedidos');
  const [busqueda, setBusqueda] = useState<string>('');
  const [filtro, setFiltro] = useState<FiltroPedidos>('todos');
  const [drawerPedidoId, setDrawerPedidoId] = useState<string | null>(null);
  const [modalEquipoAbierto, setModalEquipoAbierto] = useState<boolean>(false);
  const [modalCatalogoAbierto, setModalCatalogoAbierto] = useState<boolean>(false);
  const [toast, setToast] = useState<ToastState | null>(null);
  const [copiadoId, setCopiadoId] = useState<string | null>(null);
  const [confirmReiniciar, setConfirmReiniciar] = useState<boolean>(false);
  const [guardandoPedido, setGuardandoPedido] = useState<boolean>(false);

  const [selectedIds, setSelectedIds] = useState<Set<string>>(() => new Set<string>());
  const [lastSelectedId, setLastSelectedId] = useState<string | null>(null);
  const [batchPanel, setBatchPanel] = useState<BatchPanelActivo>(null);
  const [pendingUndoDelete, setPendingUndoDelete] = useState<PendingUndoDelete | null>(null);

  const [form, setForm] = useState<FormNuevoPedidoState>(() => crearFormInicial());
  const [lineasBorrador, setLineasBorrador] = useState<LineaPedido[]>([]);
  const [cantidadesManuales, setCantidadesManuales] = useState<Record<string, number>>({});
  const [advertenciasExcel, setAdvertenciasExcel] = useState<AdvertenciaExcel[]>([]);
  const [nombreArchivoCargado, setNombreArchivoCargado] = useState<string>('');

  const toastTimerRef = useRef<number | null>(null);
  const searchInputRef = useRef<HTMLInputElement | null>(null);
  const fileInputGeneralRef = useRef<HTMLInputElement | null>(null);
  const lastSyncRef = useRef<{ fingerprint: string }>({
    fingerprint: `${initialPedidos.length}:init:${initialPedidos.reduce((acc, p) => Math.max(acc, p.updatedAt), 0)}`
  });

  const catalogoCompleto = useMemo<ProductoCatalogo[]>(() => {
    const customMapped: ProductoCatalogo[] = productosCustom
      .filter((c) => c.activo)
      .map((c) => ({
        idGrupo: c.idGrupo,
        nombre: c.nombre,
        rubro: c.rubro,
        precioNeto: c.precioNeto,
        variantes: c.variantes,
        esCustom: true,
        activo: c.activo,
        customId: c.id
      }));
    return [...CATALOGO_MIBI, ...customMapped];
  }, [productosCustom]);

  const notify = useCallback((texto: string, tone: ToneStatus = 'neutral'): void => {
    if (toastTimerRef.current !== null) {
      window.clearTimeout(toastTimerRef.current);
    }
    setToast({ texto, tone });
    toastTimerRef.current = window.setTimeout(() => {
      setToast(null);
      toastTimerRef.current = null;
    }, 3400);
  }, []);

  const handleCopyText = useCallback(
    (key: string, text: string, msgOk: string): void => {
      copyTextSafe(text);
      setCopiadoId(key);
      notify(msgOk, 'green');
      window.setTimeout(() => setCopiadoId((curr) => (curr === key ? null : curr)), 2200);
    },
    [notify]
  );

  // Cargar filtros persistidos en navegador
  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(LS_FILTERS_KEY);
      if (!raw) return;
      const obj = JSON.parse(raw) as Record<string, unknown>;
      if (typeof obj['busqueda'] === 'string') setBusqueda(obj['busqueda']);
      const f = obj['filtro'];
      const validosDeposito = new Set([
        'todos',
        'urgentes',
        'por_armar',
        'en_armado',
        'listos',
        'despachados'
      ]);
      const validosAdminVentas = new Set([
        'todos',
        'falta_talle',
        'cheques',
        'falta_pago',
        'listos',
        'sin_enviar_deposito',
        'con_aviso_deposito'
      ]);
      if (typeof f === 'string' && (esDeposito ? validosDeposito.has(f) : validosAdminVentas.has(f))) {
        setFiltro(f as FiltroPedidos);
      }
    } catch {
      // Ignorar
    }
  }, [esDeposito]);

  useEffect(() => {
    try {
      window.localStorage.setItem(LS_FILTERS_KEY, JSON.stringify({ busqueda, filtro }));
    } catch {
      // Ignorar
    }
  }, [busqueda, filtro]);

  const refrescarCatalogoDesdeServidor = useCallback(async (): Promise<void> => {
    const res = await listarProductosCustomAction(false);
    if (res.ok && res.data) {
      setProductosCustom(res.data);
    }
  }, []);

  const refrescarCuentasDesdeServidor = useCallback(async (): Promise<void> => {
    if (!esAdmin) return;
    const res = await listarCuentasAsignablesAction();
    if (res.ok && res.data) {
      setCuentasAsignables(res.data);
    }
  }, [esAdmin]);

  const refrescarPedidosDesdeServidor = useCallback(async (): Promise<void> => {
    const res = await listarPedidosAction();
    if (res.ok && res.data) {
      setPedidos(res.data);
    }
  }, []);

  useEffect(() => {
    let active = true;
    const pollSync = async (): Promise<void> => {
      if (!active || document.visibilityState !== 'visible') return;
      try {
        const res = await fetch('/api/sync', { cache: 'no-store' });
        if (!res.ok) return;
        const data = (await res.json()) as SyncStatusResult;
        if (data.fingerprint !== lastSyncRef.current.fingerprint) {
          lastSyncRef.current.fingerprint = data.fingerprint;
          await Promise.all([refrescarPedidosDesdeServidor(), refrescarCatalogoDesdeServidor()]);
        }
      } catch {
        // Reintento en el siguiente ciclo
      }
    };

    const intervalId = window.setInterval(() => {
      void pollSync();
    }, 10000);

    const onVisibilityChange = (): void => {
      if (document.visibilityState === 'visible') {
        void pollSync();
      }
    };
    document.addEventListener('visibilitychange', onVisibilityChange);

    return () => {
      active = false;
      window.clearInterval(intervalId);
      document.removeEventListener('visibilitychange', onVisibilityChange);
    };
  }, [refrescarPedidosDesdeServidor, refrescarCatalogoDesdeServidor]);

  // Limpiar selección de IDs inexistentes
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
        if (modalCatalogoAbierto) {
          setModalCatalogoAbierto(false);
          return;
        }
        if (modalEquipoAbierto) {
          setModalEquipoAbierto(false);
          return;
        }
        if (drawerPedidoId) {
          setDrawerPedidoId(null);
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
        !esDeposito &&
        !isEditing &&
        !e.metaKey &&
        !e.ctrlKey &&
        !e.altKey &&
        e.key.toLowerCase() === 'n' &&
        !modalEquipoAbierto &&
        !modalCatalogoAbierto &&
        !drawerPedidoId
      ) {
        e.preventDefault();
        setVista('nuevo');
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [vista, esDeposito, modalEquipoAbierto, modalCatalogoAbierto, drawerPedidoId, batchPanel, selectedIds.size]);

  const resumenFinanciero = useMemo<ResumenFinanciero>(
    () => calcularResumenFinanciero(pedidos),
    [pedidos]
  );

  const resumenDeposito = useMemo(() => {
    let porArmar = 0;
    let urgentes = 0;
    let enArmado = 0;
    let listosDespachar = 0;
    for (const p of pedidos) {
      if (p.estadoArmado === 'Por armar') porArmar += 1;
      if (p.estadoArmado === 'En armado') enArmado += 1;
      if (p.estadoArmado === 'Listo' && p.estadoEnvio === 'Pendiente') listosDespachar += 1;
      const esUrg = p.asignaciones.some((a) => a.prioridad === 'URGENTE');
      if (esUrg && p.estadoEnvio === 'Pendiente') urgentes += 1;
    }
    return { porArmar, urgentes, enArmado, listosDespachar };
  }, [pedidos]);

  const conteosFiltro = useMemo(() => {
    const sinEnviarDeposito = pedidos.filter(
      (p) =>
        p.faltantesVariante === 0 &&
        p.lineas.length > 0 &&
        !p.asignaciones.some((a) => isDepositoRole(a.usuarioRol))
    ).length;
    const conAvisoDeposito = pedidos.filter((p) => Boolean(p.incidenciaDeposito)).length;

    return {
      todos: pedidos.length,
      falta_talle: pedidos.filter((p) => p.faltantesVariante > 0).length,
      cheques: pedidos.filter((p) => p.medioPago === 'Cheque' && !p.chequeAcreditado).length,
      falta_pago: pedidos.filter((p) => p.estadoPago === 'Falta pago').length,
      listos: pedidos.filter((p) => p.estadoArmado === 'Listo').length,
      sin_enviar_deposito: sinEnviarDeposito,
      con_aviso_deposito: conAvisoDeposito,
      urgentes: pedidos.filter((p) => p.asignaciones.some((a) => a.prioridad === 'URGENTE')).length,
      por_armar: pedidos.filter((p) => p.estadoArmado === 'Por armar').length,
      en_armado: pedidos.filter((p) => p.estadoArmado === 'En armado').length,
      despachados: pedidos.filter((p) => p.estadoEnvio === 'Despachado' || p.estadoEnvio === 'Entregado').length
    };
  }, [pedidos]);

  const pedidosFiltrados = useMemo<Pedido[]>(() => {
    const q = busqueda.trim().toLowerCase();
    const filtrados = pedidos.filter((p) => {
      if (filtro === 'falta_talle' && p.faltantesVariante === 0) return false;
      if (filtro === 'cheques' && !(p.medioPago === 'Cheque' && !p.chequeAcreditado)) return false;
      if (filtro === 'falta_pago' && p.estadoPago !== 'Falta pago') return false;
      if (filtro === 'listos' && p.estadoArmado !== 'Listo') return false;
      if (
        filtro === 'sin_enviar_deposito' &&
        (p.faltantesVariante > 0 ||
          p.lineas.length === 0 ||
          p.asignaciones.some((a) => isDepositoRole(a.usuarioRol)))
      ) {
        return false;
      }
      if (filtro === 'con_aviso_deposito' && !p.incidenciaDeposito) return false;
      if (filtro === 'urgentes' && !p.asignaciones.some((a) => a.prioridad === 'URGENTE')) return false;
      if (filtro === 'por_armar' && p.estadoArmado !== 'Por armar') return false;
      if (filtro === 'en_armado' && p.estadoArmado !== 'En armado') return false;
      if (filtro === 'despachados' && p.estadoEnvio !== 'Despachado' && p.estadoEnvio !== 'Entregado') {
        return false;
      }

      if (!q) return true;
      return (
        p.cliente.toLowerCase().includes(q) ||
        p.vendedor.toLowerCase().includes(q) ||
        p.numeroPedido.toLowerCase().includes(q) ||
        p.detalleResumen.toLowerCase().includes(q)
      );
    });

    if (esDeposito) {
      return [...filtrados].sort((a, b) => {
        const aUrg = a.asignaciones.some((x) => x.prioridad === 'URGENTE') ? 1 : 0;
        const bUrg = b.asignaciones.some((x) => x.prioridad === 'URGENTE') ? 1 : 0;
        if (aUrg !== bUrg) return bUrg - aUrg;
        const aRel = a.asignaciones.some((x) => x.requiereRelectura) ? 1 : 0;
        const bRel = b.asignaciones.some((x) => x.requiereRelectura) ? 1 : 0;
        if (aRel !== bRel) return bRel - aRel;
        const aDesp = a.despachoPrevisto || '9999-12-31';
        const bDesp = b.despachoPrevisto || '9999-12-31';
        if (aDesp !== bDesp) return aDesp.localeCompare(bDesp);
        return b.createdAt - a.createdAt;
      });
    }

    return filtrados;
  }, [pedidos, busqueda, filtro, esDeposito]);

  const pedidosSeleccionados = useMemo<Pedido[]>(
    () => pedidos.filter((p) => selectedIds.has(p.id)),
    [pedidos, selectedIds]
  );

  const totalSeleccionado = useMemo<number>(
    () => round2(pedidosSeleccionados.reduce((acc, p) => acc + p.totalFacturado, 0)),
    [pedidosSeleccionados]
  );

  const totalUnidadesSeleccionadas = useMemo<number>(
    () =>
      pedidosSeleccionados.reduce(
        (acc, p) => acc + p.lineas.reduce((s, l) => s + l.cantidadTotal, 0),
        0
      ),
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

  const pedidoDrawer = useMemo<Pedido | null>(
    () => (drawerPedidoId ? (pedidos.find((p) => p.id === drawerPedidoId) ?? null) : null),
    [pedidos, drawerPedidoId]
  );

  const hayBorradorSinGuardar =
    lineasBorrador.length > 0 || form.cliente.trim().length > 0 || form.observaciones.trim().length > 0;

  const proximoNumeroPedido = useMemo<string>(() => {
    const maxNum = pedidos.reduce((acc, p) => {
      const n = parseInt(p.numeroPedido.replace(/\D/g, ''), 10);
      return Number.isFinite(n) && n > acc ? n : acc;
    }, 0);
    return String(maxNum + 1).padStart(4, '0');
  }, [pedidos]);

  const handleOpenDrawer = (id: string): void => {
    setDrawerPedidoId(id);
    if (!esAdmin) {
      const target = pedidos.find((p) => p.id === id);
      const miAsig = target?.asignaciones.find((a) => a.usuarioId === session.userId);
      if (miAsig && (miAsig.vistoAt === null || miAsig.requiereRelectura)) {
        const now = Date.now();
        setPedidos((prev) =>
          prev.map((p) =>
            p.id === id
              ? {
                  ...p,
                  asignaciones: p.asignaciones.map((a) =>
                    a.usuarioId === session.userId
                      ? { ...a, vistoAt: now, requiereRelectura: false, updatedAt: now }
                      : a
                  )
                }
              : p
          )
        );
        void marcarPedidoVistoAction(id);
      }
    }
  };

  const handlePatchPedido = async (
    pedidoId: string,
    patch: Record<string, unknown>,
    motivo: string
  ): Promise<void> => {
    const res = await patchPedidoAction({ pedidoId, patch, motivo });
    if (!res.ok || !res.data) {
      notify(res.error ?? 'No se pudo actualizar el pedido.', 'red');
      return;
    }
    const updated = res.data;
    setPedidos((prev) => prev.map((p) => (p.id === updated.id ? updated : p)));
  };

  const handleConfirmarAcreditacion = (pedido: Pedido): void => {
    if (!esAdmin) {
      notify('Solo Administración puede confirmar acreditación de cheques.', 'amber');
      return;
    }
    const hoy = getTodayISO();
    const nota = `Cheque acreditado el ${formatShortDate(hoy)}`;
    void handlePatchPedido(
      pedido.id,
      {
        medioPago: 'Transferencia / Efectivo',
        chequeAcreditado: true,
        montoAbonado: pedido.totalFacturado,
        montoCheque: 0,
        observaciones: [pedido.observaciones, nota].filter(Boolean).join(' · ')
      },
      `Confirmó plata acreditada (${formatARS(pedido.totalFacturado)}) → Pagado`
    ).then(() => {
      notify(`Pedido #${pedido.numeroPedido} pasado a Pagado.`, 'green');
    });
  };

  const handleResolverEnDrawer = (pedido: Pedido, idLinea: string, cantidades: number[]): void => {
    const lineaObj = pedido.lineas.find((l) => l.idLinea === idLinea);
    if (!lineaObj) return;
    const resuelta = aplicarResolucionLinea(lineaObj, cantidades, catalogoCompleto);
    if (!resuelta) {
      notify(`La suma debe ser exactamente ${lineaObj.cantidadTotal} u.`, 'red');
      return;
    }
    const nuevasLineas = pedido.lineas.map((l) => (l.idLinea === idLinea ? resuelta : l));
    void handlePatchPedido(
      pedido.id,
      { lineas: nuevasLineas },
      `Definió talle/color en ${lineaObj.nombreProducto} (${cantidades.join(' / ')})`
    ).then(() => {
      notify(`${lineaObj.nombreProducto} confirmado para depósito.`, 'green');
    });
  };

  const handleEnviarACuentas = async (params: {
    pedidoIds: string[];
    usuarioIds: string[];
    prioridad: PrioridadAsignacion;
    notaEnvio: string;
    forzarSinPago?: boolean;
  }): Promise<EnviarCuentasResult> => {
    const res = await enviarPedidosACuentasAction(params);
    if (res.ok) {
      await refrescarPedidosDesdeServidor();
      if (params.pedidoIds.length > 1) {
        const omitidos = [...res.omitidosTalle, ...res.omitidosSinLineas];
        if (omitidos.length > 0) {
          notify(
            `Enviados ${res.enviados} pedido(s) · Omitidos por faltante: ${omitidos.join(', ')}`,
            'amber'
          );
        } else {
          notify(`${res.enviados} pedido(s) enviados a la cuenta seleccionada.`, 'green');
        }
      }
    } else if (!res.requiereConfirmacionSinPago && params.pedidoIds.length > 1) {
      notify(res.error ?? 'No se pudieron enviar los pedidos.', 'red');
    }
    return res;
  };

  const handleRetirarAsignacion = async (pedidoId: string, usuarioId: string): Promise<void> => {
    const res = await retirarAsignacionAction({ pedidoId, usuarioId });
    if (!res.ok) {
      notify(res.error ?? 'No se pudo retirar la asignación.', 'red');
      return;
    }
    await refrescarPedidosDesdeServidor();
    notify('Asignación retirada.', 'neutral');
  };

  const handleReportarIncidencia = async (pedidoId: string, mensaje: string): Promise<void> => {
    const res = await reportarIncidenciaDepositoAction({ pedidoId, mensaje });
    if (!res.ok || !res.data) {
      notify(res.error ?? 'No se pudo enviar el aviso.', 'red');
      return;
    }
    const updated = res.data;
    setPedidos((prev) => prev.map((p) => (p.id === updated.id ? updated : p)));
    notify('Aviso enviado a Administración.', 'amber');
  };

  const handleResolverIncidencia = async (pedidoId: string): Promise<void> => {
    const res = await resolverIncidenciaAction(pedidoId);
    if (!res.ok || !res.data) {
      notify(res.error ?? 'No se pudo marcar como resuelto.', 'red');
      return;
    }
    const updated = res.data;
    setPedidos((prev) => prev.map((p) => (p.id === updated.id ? updated : p)));
    notify('Aviso marcado como resuelto.', 'green');
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
            const row = pedidosFiltrados[i];
            if (row) {
              if (marcar) next.add(row.id);
              else next.delete(row.id);
            }
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

  const handleBatchArmado = (destino: 'En armado' | 'Listo'): void => {
    const ids = pedidosSeleccionados.map((p) => p.id);
    setBatchPanel(null);
    void batchPedidosAction({ ids, action: { type: 'armado', estadoArmado: destino } }).then(async (res) => {
      if (!res.ok || !res.data) {
        notify(res.error ?? 'Error al actualizar lote.', 'red');
        return;
      }
      await refrescarPedidosDesdeServidor();
      const { actualizados, omitidos } = res.data;
      if (omitidos.length > 0 && actualizados > 0) {
        notify(
          `${actualizados} pasados a Listo · ${omitidos.length} bloqueados por falta de talle (${omitidos.join(', ')})`,
          'amber'
        );
      } else if (omitidos.length > 0 && actualizados === 0) {
        notify(`Ninguno pudo pasar a Listo: falta definir talle o color en ${omitidos.join(', ')}`, 'red');
      } else {
        notify(`${actualizados} ${actualizados === 1 ? 'pedido pasado' : 'pedidos pasados'} a ${destino}`, 'green');
      }
    });
  };

  const handleBatchEnvio = (destino: 'Despachado' | 'Entregado', transporte: string): void => {
    const ids = pedidosSeleccionados.map((p) => p.id);
    const transp = transporte.trim();
    setBatchPanel(null);
    void batchPedidosAction({
      ids,
      action: { type: 'envio', estadoEnvio: destino, ...(transp ? { transporte: transp } : {}) }
    }).then(async (res) => {
      if (!res.ok || !res.data) {
        notify(res.error ?? 'Error al actualizar envíos.', 'red');
        return;
      }
      await refrescarPedidosDesdeServidor();
      notify(
        transp
          ? `${res.data.actualizados} pedidos marcados como ${destino} (${transp})`
          : `${res.data.actualizados} pedidos marcados como ${destino}`,
        'green'
      );
    });
  };

  const handleBatchPago = (accion: 'Pagado' | 'Falta pago'): void => {
    if (!esAdmin) {
      notify('Solo Administración puede modificar pagos en lote.', 'amber');
      return;
    }
    const ids = pedidosSeleccionados.map((p) => p.id);
    setBatchPanel(null);
    void batchPedidosAction({ ids, action: { type: 'pago', estado: accion } }).then(async (res) => {
      if (!res.ok || !res.data) {
        notify(res.error ?? 'Error al actualizar pagos.', 'red');
        return;
      }
      await refrescarPedidosDesdeServidor();
      notify(
        `${res.data.actualizados} ${res.data.actualizados === 1 ? 'pedido actualizado' : 'pedidos actualizados'} a ${accion}`,
        accion === 'Pagado' ? 'green' : 'neutral'
      );
    });
  };

  const handleBatchVendedor = (vendedor: string): void => {
    if (!esAdmin) {
      notify('Solo Administración puede reasignar vendedores en lote.', 'amber');
      return;
    }
    const ids = pedidosSeleccionados.map((p) => p.id);
    setBatchPanel(null);
    void batchPedidosAction({ ids, action: { type: 'vendedor', vendedor } }).then(async (res) => {
      if (!res.ok || !res.data) {
        notify(res.error ?? 'Error al asignar vendedor.', 'red');
        return;
      }
      await refrescarPedidosDesdeServidor();
      notify(`${res.data.actualizados} pedidos asignados a ${vendedor}`, 'green');
    });
  };

  const handleEliminarIndividual = (pedido: Pedido): void => {
    if (!esAdmin) {
      notify('Solo Administración puede eliminar pedidos.', 'amber');
      return;
    }
    void softDeletePedidosAction({ ids: [pedido.id] }).then(async (res) => {
      if (!res.ok) {
        notify(res.error ?? 'No se pudo eliminar.', 'red');
        return;
      }
      if (drawerPedidoId === pedido.id) setDrawerPedidoId(null);
      await refrescarPedidosDesdeServidor();
      notify(`Pedido #${pedido.numeroPedido} eliminado.`, 'neutral');
    });
  };

  const handleConfirmarEliminarLote = (): void => {
    if (!esAdmin || pedidosSeleccionados.length === 0) return;
    if (pendingUndoDelete) {
      window.clearTimeout(pendingUndoDelete.timerId);
    }
    const aEliminar = [...pedidosSeleccionados];
    const ids = aEliminar.map((p) => p.id);
    limpiarSeleccion();

    void softDeletePedidosAction({ ids }).then(async (res) => {
      if (!res.ok) {
        notify(res.error ?? 'No se pudieron eliminar los pedidos.', 'red');
        return;
      }
      if (drawerPedidoId && ids.includes(drawerPedidoId)) setDrawerPedidoId(null);
      await refrescarPedidosDesdeServidor();

      const timerId = window.setTimeout(() => {
        setPendingUndoDelete((curr) => (curr && curr.timerId === timerId ? null : curr));
      }, 6000);
      setPendingUndoDelete({ pedidos: aEliminar, timerId });
    });
  };

  const handleDeshacerEliminacion = (): void => {
    if (!pendingUndoDelete) return;
    window.clearTimeout(pendingUndoDelete.timerId);
    const ids = pendingUndoDelete.pedidos.map((p) => p.id);
    setPendingUndoDelete(null);

    void restorePedidosAction({ ids }).then(async (res) => {
      if (!res.ok) {
        notify(res.error ?? 'No se pudieron restaurar.', 'red');
        return;
      }
      await refrescarPedidosDesdeServidor();
      notify(
        `${ids.length} ${ids.length === 1 ? 'pedido restaurado' : 'pedidos restaurados'}.`,
        'green'
      );
    });
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
    setDrawerPedidoId(null);
    setVista('nuevo');
    notify(`Datos de ${pedido.cliente} copiados para un nuevo pedido.`, 'green');
  };

  const handleGuardarNuevoPedido = async (payload: {
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
  }): Promise<void> => {
    setGuardandoPedido(true);
    try {
      const res = await crearPedidoAction(payload);
      if (!res.ok || !res.data) {
        notify(res.error ?? 'No se pudo guardar el pedido.', 'red');
        return;
      }
      const nuevo = res.data;
      setPedidos((prev) => [nuevo, ...prev]);
      setForm(crearFormInicial());
      setLineasBorrador([]);
      setAdvertenciasExcel([]);
      setCantidadesManuales({});
      setNombreArchivoCargado('');
      setVista('pedidos');
      notify(`Pedido #${nuevo.numeroPedido} guardado (${formatARS(nuevo.totalFacturado)}).`, 'green');
    } finally {
      setGuardandoPedido(false);
    }
  };

  const handleImportarPlanillaGeneral = async (file: File): Promise<void> => {
    if (!esAdmin) {
      notify('Solo Administración puede importar planillas previas.', 'amber');
      return;
    }
    const fd = new FormData();
    fd.append('file', file);
    const res = await importarPlanillaAction(fd);
    if (!res.ok || !res.data) {
      notify(res.error ?? 'Error al importar planilla.', 'red');
      return;
    }
    if (res.data.importados === 0) {
      notify(
        res.data.duplicadosOmitidos > 0
          ? `Sin pedidos nuevos (${res.data.duplicadosOmitidos} duplicados omitidos).`
          : 'No se encontraron pedidos válidos en la planilla.',
        'amber'
      );
      return;
    }
    await refrescarPedidosDesdeServidor();
    notify(
      `Importados ${res.data.importados} pedidos (${res.data.duplicadosOmitidos} duplicados omitidos).`,
      'green'
    );
  };

  const handleExportarHojaDeposito = async (pedido: Pedido): Promise<void> => {
    const res = await fetch(`/api/export/deposito/${encodeURIComponent(pedido.id)}`);
    if (!res.ok) {
      const body = (await res.json().catch(() => ({}))) as { error?: string };
      notify(body.error ?? 'No se pudo exportar este pedido.', 'red');
      return;
    }
    await descargarResponseBlob(res, `Mibi_Deposito_Pedido_${pedido.numeroPedido}.xlsx`);
    notify(`Descargada Hoja depósito #${pedido.numeroPedido}`, 'green');
  };

  const handleExportarTablero = async (ids?: string[]): Promise<void> => {
    if (!esAdmin) {
      notify('Solo Administración puede exportar el tablero general.', 'amber');
      return;
    }
    const res = await fetch('/api/export/tablero', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(ids && ids.length > 0 ? { ids } : {})
    });
    if (!res.ok) {
      const body = (await res.json().catch(() => ({}))) as { error?: string };
      notify(body.error ?? 'No se pudo exportar Excel.', 'red');
      return;
    }
    await descargarResponseBlob(res, `PEDIDOS_MAYORISTA_MIBI_${getTodayISO()}.xlsx`);
    const cant = ids && ids.length > 0 ? ids.length : pedidos.length;
    notify(`Exportado PEDIDOS_MAYORISTA_MIBI (${cant} pedidos).`, 'green');
  };

  const handleExportarPicking = async (): Promise<void> => {
    const ids = pedidosSeleccionados.map((p) => p.id);
    setBatchPanel(null);
    const res = await fetch('/api/export/picking', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ids })
    });
    if (!res.ok) {
      const body = (await res.json().catch(() => ({}))) as { error?: string };
      notify(body.error ?? 'No se pudo generar el consolidado.', 'red');
      return;
    }

    const incluidos = Number(res.headers.get('X-Mibi-Incluidos') ?? '0');
    const unidades = Number(res.headers.get('X-Mibi-Unidades') ?? '0');
    const exclTalle = JSON.parse(
      decodeURIComponent(res.headers.get('X-Mibi-Excluidos-Talle') ?? '[]')
    ) as string[];
    const exclSin = JSON.parse(
      decodeURIComponent(res.headers.get('X-Mibi-Excluidos-Sin-Lineas') ?? '[]')
    ) as string[];

    await descargarResponseBlob(res, `Mibi_Picking_Consolidado_${incluidos}ped_${getTodayISO()}.xlsx`);

    const notas: string[] = [];
    if (exclTalle.length > 0) notas.push(`excluidos sin talle: ${exclTalle.join(', ')}`);
    if (exclSin.length > 0) notas.push(`sin códigos: ${exclSin.join(', ')}`);
    notify(
      notas.length > 0
        ? `Consolidado listo (${incluidos} pedidos · ${unidades} u. · ${notas.join(' · ')})`
        : `Consolidado para Depósito listo (${incluidos} pedidos · ${unidades} u.).`,
      notas.length > 0 ? 'amber' : 'green'
    );
  };

  const handleReiniciarSistema = async (): Promise<void> => {
    if (!esAdmin) return;
    if (!confirmReiniciar) {
      setConfirmReiniciar(true);
      window.setTimeout(() => setConfirmReiniciar(false), 4000);
      return;
    }
    const res = await reiniciarSistemaAction();
    if (!res.ok) {
      notify(res.error ?? 'No se pudo reiniciar el sistema.', 'red');
      return;
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
    notify('Sistema reiniciado en cero.', 'neutral');
  };

  const filtrosPildora: ReadonlyArray<{ id: FiltroPedidos; label: string; count: number }> = esDeposito
    ? [
        { id: 'todos', label: 'Todos', count: conteosFiltro.todos },
        { id: 'urgentes', label: 'Urgentes', count: conteosFiltro.urgentes },
        { id: 'por_armar', label: 'Por armar', count: conteosFiltro.por_armar },
        { id: 'en_armado', label: 'En armado', count: conteosFiltro.en_armado },
        { id: 'listos', label: 'Listos', count: conteosFiltro.listos },
        { id: 'despachados', label: 'Despachados', count: conteosFiltro.despachados }
      ]
    : [
        { id: 'todos', label: 'Todos', count: conteosFiltro.todos },
        { id: 'falta_talle', label: 'Falta talle', count: conteosFiltro.falta_talle },
        { id: 'cheques', label: 'Cheques', count: conteosFiltro.cheques },
        { id: 'falta_pago', label: 'Falta pago', count: conteosFiltro.falta_pago },
        { id: 'listos', label: 'Listos', count: conteosFiltro.listos },
        ...(esAdmin
          ? [
              {
                id: 'sin_enviar_deposito' as const,
                label: 'Sin enviar a depósito',
                count: conteosFiltro.sin_enviar_deposito
              }
            ]
          : []),
        ...(esAdmin && conteosFiltro.con_aviso_deposito > 0
          ? [
              {
                id: 'con_aviso_deposito' as const,
                label: 'Con aviso de depósito',
                count: conteosFiltro.con_aviso_deposito
              }
            ]
          : [])
      ];

  return (
    <div className="min-h-screen bg-[#FAFAFA] text-[#111111] flex flex-col">
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

      {pendingUndoDelete && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 z-50">
          <div className="flex items-center gap-3 pl-4 pr-2 py-2 rounded-full bg-zinc-900 text-white text-xs font-medium shadow-lg border border-zinc-800">
            <span>
              {pendingUndoDelete.pedidos.length === 1
                ? `Pedido #${pendingUndoDelete.pedidos[0]?.numeroPedido} eliminado`
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

      <div className="h-[2px] w-full bg-[#F4B6CE]" />
      <header className="sticky top-0 z-30 h-[60px] bg-white/95 backdrop-blur border-b border-[#EAEAEA]">
        <div className="max-w-7xl mx-auto h-full px-4 sm:px-6 flex items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="inline-flex items-center">
              <MibiLogo width={64} variant="pink" />
            </div>
            <span className="hidden md:inline-flex items-center gap-1.5 text-[11px] text-zinc-400">
              <span className={`w-1.5 h-1.5 rounded-full ${isLocalDev ? 'bg-amber-400' : 'bg-emerald-500'}`} />
              {isLocalDev ? 'Entorno Local' : 'Nube privada'}
            </span>
          </div>

          {!esDeposito && (
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
          )}

          <div className="flex items-center gap-2">
            {!esDeposito && (
              <button
                type="button"
                onClick={() => setModalCatalogoAbierto(true)}
                className="hidden sm:inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors"
                title="Productos fuera de planilla"
              >
                <BookOpen className="w-3.5 h-3.5" strokeWidth={1.5} />
                Catálogo
              </button>
            )}

            {esAdmin && pedidos.length > 0 && (
              <button
                type="button"
                onClick={() => void handleExportarTablero()}
                className="hidden sm:inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors"
              >
                <Download className="w-3.5 h-3.5" strokeWidth={1.5} />
                Exportar Excel
              </button>
            )}

            {esAdmin && (
              <button
                type="button"
                onClick={() => fileInputGeneralRef.current?.click()}
                className="hidden lg:inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors"
                title="Importar planilla general previa PEDIDOS MAYORISTA MIBI.xlsx"
              >
                <Upload className="w-3.5 h-3.5" strokeWidth={1.5} />
                Importar planilla
              </button>
            )}

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

            {esDeposito ? (
              <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-zinc-50 border border-zinc-200/80 text-zinc-700">
                <Users className="w-3.5 h-3.5 text-zinc-500" strokeWidth={1.5} />
                <span>{session.nombre}</span>
              </span>
            ) : (
              <button
                type="button"
                onClick={() => setModalEquipoAbierto(true)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-zinc-50 border border-zinc-200/80 text-zinc-700 hover:bg-zinc-100 transition-colors"
                title="Equipo y verificación técnica"
              >
                <Users className="w-3.5 h-3.5 text-zinc-500" strokeWidth={1.5} />
                <span className="hidden sm:inline">{session.nombre}</span>
              </button>
            )}

            <button
              type="button"
              onClick={() => void logoutAction()}
              className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-800 hover:bg-zinc-100 transition-colors"
              title="Bloquear sesión"
              aria-label="Bloquear sesión"
            >
              <LogOut className="w-4 h-4" strokeWidth={1.5} />
            </button>
          </div>
        </div>
      </header>

      <main className="flex-1 w-full max-w-7xl mx-auto px-4 sm:px-6 py-7">
        {(vista === 'pedidos' || esDeposito) && (
          <div className="space-y-6">
            {esDeposito ? (
              <section className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-5 py-4">
                  <div className="text-xs font-medium text-zinc-500">Por armar</div>
                  <div className="mt-1.5 text-2xl font-semibold tracking-tight text-[#111111] tabular-nums">
                    {resumenDeposito.porArmar}
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-5 py-4">
                  <div className="text-xs font-medium text-zinc-500">Urgentes</div>
                  <div className="mt-1.5 text-2xl font-semibold tracking-tight text-red-700 tabular-nums">
                    {resumenDeposito.urgentes}
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-5 py-4">
                  <div className="text-xs font-medium text-zinc-500">En armado</div>
                  <div className="mt-1.5 text-2xl font-semibold tracking-tight text-[#111111] tabular-nums">
                    {resumenDeposito.enArmado}
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-5 py-4">
                  <div className="text-xs font-medium text-zinc-500">Listos para despachar</div>
                  <div className="mt-1.5 text-2xl font-semibold tracking-tight text-emerald-700 tabular-nums">
                    {resumenDeposito.listosDespachar}
                  </div>
                </div>
              </section>
            ) : (
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
            )}

            {pedidos.length === 0 ? (
              <section className="bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] px-6 py-16 text-center space-y-4">
                <div className="mx-auto w-11 h-11 rounded-full bg-zinc-100 flex items-center justify-center text-zinc-500">
                  <Package className="w-5 h-5" strokeWidth={1.5} />
                </div>
                {esDeposito ? (
                  <div className="space-y-1">
                    <h2 className="text-base font-semibold text-[#111111]">
                      No hay pedidos asignados a Depósito todavía
                    </h2>
                    <p className="text-xs text-zinc-500 max-w-md mx-auto">
                      Cuando Administración envíe un pedido listo a tu cuenta, aparecerá acá automáticamente.
                    </p>
                  </div>
                ) : (
                  <>
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
                      {esAdmin && (
                        <button
                          type="button"
                          onClick={() => fileInputGeneralRef.current?.click()}
                          className="inline-flex items-center gap-1.5 px-4 py-2.5 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors"
                        >
                          <Upload className="w-3.5 h-3.5" strokeWidth={1.5} />
                          Importar planilla previa
                        </button>
                      )}
                    </div>
                  </>
                )}
              </section>
            ) : (
              <section className="space-y-4">
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
                      placeholder="Buscar cliente, número o artículo..."
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
                    {filtrosPildora.map((item) => {
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
                          <span className={`text-[11px] tabular-nums ${activo ? 'text-zinc-300' : 'text-zinc-400'}`}>
                            {item.count}
                          </span>
                        </button>
                      );
                    })}

                    {!esDeposito && (
                      <button
                        type="button"
                        onClick={() => setVista('nuevo')}
                        className="ml-1 inline-flex items-center gap-1.5 px-4 py-1.5 rounded-full text-xs font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors"
                      >
                        <Plus className="w-3.5 h-3.5" strokeWidth={1.75} />
                        Nuevo pedido
                      </button>
                    )}
                  </div>
                </div>

                <PedidosTable
                  pedidosFiltrados={pedidosFiltrados}
                  selectedIds={selectedIds}
                  todosVisiblesSeleccionados={todosVisiblesSeleccionados}
                  seleccionParcialVisible={seleccionParcialVisible}
                  esAdmin={esAdmin}
                  esDeposito={esDeposito}
                  onToggleSelectAll={handleToggleSelectAllVisible}
                  onToggleSelectRow={handleToggleSelectRow}
                  onOpenDrawer={handleOpenDrawer}
                  onConfirmarAcreditacion={handleConfirmarAcreditacion}
                  onExportarHojaDeposito={(p) => void handleExportarHojaDeposito(p)}
                />
              </section>
            )}
          </div>
        )}

        {!esDeposito && vista === 'nuevo' && (
          <NuevoPedidoView
            esAdmin={esAdmin}
            proximoNumeroPedido={proximoNumeroPedido}
            catalogoCompleto={catalogoCompleto}
            form={form}
            setForm={setForm}
            lineasBorrador={lineasBorrador}
            setLineasBorrador={setLineasBorrador}
            cantidadesManuales={cantidadesManuales}
            setCantidadesManuales={setCantidadesManuales}
            advertenciasExcel={advertenciasExcel}
            setAdvertenciasExcel={setAdvertenciasExcel}
            nombreArchivoCargado={nombreArchivoCargado}
            setNombreArchivoCargado={setNombreArchivoCargado}
            copiadoId={copiadoId}
            guardandoPedido={guardandoPedido}
            onVolver={() => setVista('pedidos')}
            onCopyText={handleCopyText}
            onNotify={notify}
            onProductoCustomCreado={(rec) => setProductosCustom((prev) => [rec, ...prev])}
            onGuardar={handleGuardarNuevoPedido}
          />
        )}
      </main>

      <footer className="py-6 text-center text-[11px] text-zinc-400 border-t border-zinc-200/50">
        {MIBI_BRAND.pieOficial}
      </footer>

      {vista === 'pedidos' && selectedIds.size > 0 && (
        <FloatingBatchBar
          selectedCount={selectedIds.size}
          totalSeleccionado={totalSeleccionado}
          totalUnidadesSeleccionadas={totalUnidadesSeleccionadas}
          pedidosSeleccionados={pedidosSeleccionados}
          esAdmin={esAdmin}
          esDeposito={esDeposito}
          cuentasAsignables={cuentasAsignables}
          batchPanel={batchPanel}
          setBatchPanel={setBatchPanel}
          onLimpiarSeleccion={limpiarSeleccion}
          onBatchArmado={handleBatchArmado}
          onBatchEnvio={handleBatchEnvio}
          onBatchPago={handleBatchPago}
          onBatchVendedor={handleBatchVendedor}
          onEnviarLoteACuentas={handleEnviarACuentas}
          onExportTableroSeleccionados={() => {
            setBatchPanel(null);
            void handleExportarTablero(pedidosSeleccionados.map((p) => p.id));
          }}
          onExportPickingConsolidado={() => void handleExportarPicking()}
          onConfirmarEliminarLote={handleConfirmarEliminarLote}
        />
      )}

      {pedidoDrawer && (
        <PedidoDrawer
          pedido={pedidoDrawer}
          esAdmin={esAdmin}
          esDeposito={esDeposito}
          catalogoCompleto={catalogoCompleto}
          cuentasAsignables={cuentasAsignables}
          copiadoId={copiadoId}
          onClose={() => setDrawerPedidoId(null)}
          onPatch={handlePatchPedido}
          onConfirmarAcreditacion={handleConfirmarAcreditacion}
          onResolverLinea={handleResolverEnDrawer}
          onExportarHojaDeposito={(p) => void handleExportarHojaDeposito(p)}
          onDuplicarComoNuevo={handleDuplicarComoNuevoPedido}
          onEliminarIndividual={handleEliminarIndividual}
          onEnviarACuentas={handleEnviarACuentas}
          onRetirarAsignacion={handleRetirarAsignacion}
          onReportarIncidencia={handleReportarIncidencia}
          onResolverIncidencia={handleResolverIncidencia}
          onCopyText={handleCopyText}
          onNotify={notify}
        />
      )}

      {modalEquipoAbierto && !esDeposito && (
        <EquipoModal
          session={session}
          onClose={() => setModalEquipoAbierto(false)}
          onNotify={notify}
          onEquipoActualizado={refrescarCuentasDesdeServidor}
        />
      )}

      {modalCatalogoAbierto && !esDeposito && (
        <CatalogoModal
          esAdmin={esAdmin}
          onClose={() => setModalCatalogoAbierto(false)}
          onCatalogoActualizado={refrescarCatalogoDesdeServidor}
          onNotify={notify}
        />
      )}
    </div>
  );
}
