'use client';

import React, { useEffect, useState } from 'react';
import {
  AlertTriangle,
  Check,
  Clock,
  Copy,
  Download,
  Lock,
  MessageCircle,
  Printer,
  Send,
  Trash2,
  UserCheck,
  X
} from 'lucide-react';
import type { EnviarCuentasResult } from '@/lib/dal/pedidos';
import { formatARS, formatShortDate, getTodayISO, round2 } from '@/lib/finanzas';
import {
  isDepositoRole,
  type EstadoArmado,
  type EstadoEnvio,
  type Pedido,
  type PrioridadAsignacion,
  type ProductoCatalogo,
  type ToneStatus,
  type UsuarioEquipo
} from '@/lib/types';
import { generarMensajeDespachoWhatsApp, generarMensajeWhatsApp } from '@/lib/whatsapp';
import { MibiLogo } from './mibi-logo';
import { StatusPill } from './status-pill';

interface DrawerDraft {
  cliente: string;
  whatsapp: string;
  descuentoPorc: string;
  montoAbonado: string;
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
  copiadoId: string | null;
  onClose: () => void;
  onPatch: (
    pedidoId: string,
    patch: Record<string, unknown>,
    motivo: string
  ) => Promise<void>;
  onConfirmarAcreditacion: (pedido: Pedido) => void;
  onResolverLinea: (pedido: Pedido, idLinea: string, cantidades: number[]) => void;
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

export function PedidoDrawer({
  pedido,
  esAdmin,
  esDeposito,
  catalogoCompleto,
  cuentasAsignables,
  copiadoId,
  onClose,
  onPatch,
  onConfirmarAcreditacion,
  onResolverLinea,
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
  const [mostrarHojaImprimible, setMostrarHojaImprimible] = useState<boolean>(false);
  const [confirmEliminar, setConfirmEliminar] = useState<boolean>(false);
  const [resolucionesManual, setResolucionesManual] = useState<Record<string, string[]>>({});
  const [draft, setDraft] = useState<DrawerDraft>({
    cliente: pedido.cliente,
    whatsapp: pedido.whatsapp,
    descuentoPorc: String(pedido.descuentoPorc),
    montoAbonado: String(pedido.montoAbonado),
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

  useEffect(() => {
    setConfirmEliminar(false);
    setMostrarHojaImprimible(false);
    setMostrarEnviar(false);
    setRequiereConfirmSinPago(false);
    setMostrarIncidenciaForm(false);
    setIncidenciaTexto('');
  }, [pedido.id]);

  useEffect(() => {
    setDraft({
      cliente: pedido.cliente,
      whatsapp: pedido.whatsapp,
      descuentoPorc: String(pedido.descuentoPorc),
      montoAbonado: String(pedido.montoAbonado),
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
    pedido.fechaAcreditacionCheque,
    pedido.detalleCheque,
    pedido.transporte,
    pedido.nroSeguimiento,
    pedido.observaciones
  ]);

  const commitCliente = (): void => {
    const val = draft.cliente.trim();
    if (!val) {
      setDraft((prev) => ({ ...prev, cliente: pedido.cliente }));
      onNotify('El nombre del cliente no puede quedar vacío.', 'amber');
      return;
    }
    if (val !== pedido.cliente) {
      const motivo = `Actualizó cliente: ${pedido.cliente} → ${val}`.slice(0, 300);
      void onPatch(pedido.id, { cliente: val }, motivo).then(() => {
        onNotify('Cliente actualizado.', 'green');
      });
    }
  };

  const commitWhatsapp = (): void => {
    const val = draft.whatsapp.trim();
    if (val !== pedido.whatsapp) {
      void onPatch(pedido.id, { whatsapp: val }, `Actualizó WhatsApp: ${val || 'vacío'}`);
    }
  };

  const handleConfirmarEnvioCuenta = async (forzarSinPago = false): Promise<void> => {
    if (!cuentaDestinoId) {
      onNotify('Elegí una cuenta de destino.', 'amber');
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

  const asignacionesDeposito = pedido.asignaciones.filter((a) => isDepositoRole(a.usuarioRol));
  const esUrgenteDeposito = asignacionesDeposito.some((a) => a.prioridad === 'URGENTE');
  const notasDeposito = Array.from(
    new Set(asignacionesDeposito.map((a) => a.notaEnvio.trim()).filter(Boolean))
  ).join(' · ');
  const requiereRelectura = asignacionesDeposito.some((a) => a.requiereRelectura);
  const totalUnidades = pedido.lineas.reduce((acc, l) => acc + l.cantidadTotal, 0);

  const cuentasDeposito = cuentasAsignables.filter((c) => isDepositoRole(c.rol));
  const cuentasVentas = cuentasAsignables.filter((c) => !isDepositoRole(c.rol));

  return (
    <div
      className="fixed inset-0 z-50 flex justify-end bg-black/25 backdrop-blur-[1px]"
      onClick={onClose}
    >
      <aside
        onClick={(e) => e.stopPropagation()}
        className="w-full max-w-xl bg-white h-full overflow-y-auto border-l border-[#EAEAEA] shadow-2xl flex flex-col"
      >
        <div className="sticky top-0 z-10 bg-white/95 backdrop-blur border-b border-[#EAEAEA] px-6 py-4 flex items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h2 className="text-base font-semibold text-[#111111] tabular-nums">
                Pedido #{pedido.numeroPedido}
              </h2>
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
            <p className="text-xs text-zinc-500 mt-0.5">
              {pedido.cliente} · {pedido.vendedor}
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-2 rounded-full text-zinc-400 hover:text-zinc-900 hover:bg-zinc-100"
            aria-label="Cerrar panel"
          >
            <X className="w-4 h-4" strokeWidth={1.5} />
          </button>
        </div>

        <div className="p-6 space-y-6 flex-1">
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              disabled={pedido.faltantesVariante > 0 || pedido.lineas.length === 0}
              onClick={() => onExportarHojaDeposito(pedido)}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
            >
              <Download className="w-3.5 h-3.5" strokeWidth={1.5} />
              Hoja depósito (.xlsx)
            </button>

            <button
              type="button"
              disabled={pedido.faltantesVariante > 0}
              onClick={() => setMostrarHojaImprimible(!mostrarHojaImprimible)}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
            >
              <Printer className="w-3.5 h-3.5" strokeWidth={1.5} />
              {mostrarHojaImprimible ? 'Ver gestión' : 'Vista imprimible'}
            </button>

            {!esDeposito && (
              <button
                type="button"
                onClick={() => onDuplicarComoNuevo(pedido)}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-medium bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 transition-colors"
                title="Crea un pedido nuevo prellenando cliente, vendedor y WhatsApp"
              >
                <Copy className="w-3.5 h-3.5" strokeWidth={1.5} />
                Duplicar como nuevo
              </button>
            )}
          </div>

          {mostrarHojaImprimible ? (
            <div className="mibi-print-area space-y-4 border border-zinc-200 rounded-2xl p-5">
              <div className="flex items-start justify-between border-b border-zinc-200 pb-3">
                <div>
                  <MibiLogo width={76} variant="pink" />
                  <div className="text-xs font-semibold mt-2 flex items-center gap-2">
                    <span>HOJA PARA DEPÓSITO — PEDIDO #{pedido.numeroPedido}</span>
                    {esUrgenteDeposito && (
                      <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-red-100 text-red-700">
                        URGENTE
                      </span>
                    )}
                  </div>
                  <div className="text-[11px] text-zinc-500">
                    Cliente: {pedido.cliente} · Transporte: {pedido.transporte || 'A definir'}
                    {pedido.despachoPrevisto ? ` · Despacho: ${formatShortDate(pedido.despachoPrevisto)}` : ''}
                  </div>
                  {notasDeposito && (
                    <div className="text-[11px] text-zinc-700 mt-1">
                      <strong>Nota operativa:</strong> {notasDeposito}
                    </div>
                  )}
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
                    <th className="py-2 text-left">Producto</th>
                    <th className="py-2 text-left">Talle / Color</th>
                    <th className="py-2 text-right">Cantidad pedida</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-200/60 tabular-nums">
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
                </tbody>
                <tfoot>
                  <tr className="border-t border-zinc-300 font-semibold">
                    <td colSpan={3} className="py-2.5 text-right text-zinc-600">
                      TOTAL UNIDADES
                    </td>
                    <td className="py-2.5 text-right tabular-nums">{totalUnidades}</td>
                  </tr>
                </tfoot>
              </table>
              <p className="text-[10px] text-zinc-400 pt-2 border-t border-zinc-200">
                Mibi® · Ciudad de la Paz 2467, Belgrano, CABA · WhatsApp 11-5606-9384 · contacto@mibi.com.ar
              </p>
            </div>
          ) : (
            <>
              {esDeposito && (notasDeposito || esUrgenteDeposito || requiereRelectura) && (
                <div
                  className={`rounded-2xl border p-4 space-y-1.5 text-xs ${
                    esUrgenteDeposito
                      ? 'bg-red-50/70 border-red-200 text-red-900'
                      : 'bg-amber-50/70 border-amber-200 text-amber-900'
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
                <div className="bg-amber-50/80 border border-amber-200 rounded-2xl p-4 space-y-2 text-xs">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-amber-900 flex items-center gap-1.5">
                      <AlertTriangle className="w-4 h-4 text-amber-600" strokeWidth={1.75} />
                      Aviso de Depósito {pedido.incidenciaPor ? `(${pedido.incidenciaPor})` : ''}
                    </span>
                    {esAdmin && (
                      <button
                        type="button"
                        onClick={() => void onResolverIncidencia(pedido.id)}
                        className="px-3 py-1 rounded-full text-[11px] font-medium bg-white border border-amber-300 text-amber-900 hover:bg-amber-100"
                      >
                        Marcar resuelto
                      </button>
                    )}
                  </div>
                  <p className="text-amber-900">{pedido.incidenciaDeposito}</p>
                </div>
              )}

              {!esDeposito && pedido.faltantesVariante > 0 && (
                <div className="bg-red-50/70 border border-red-200 rounded-2xl p-4 space-y-3">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs font-semibold text-red-800">
                      Falta definir talle o color ({pedido.faltantesVariante})
                    </span>
                    <button
                      type="button"
                      onClick={() => {
                        const msg = generarMensajeWhatsApp(
                          pedido.cliente,
                          pedido.numeroPedido,
                          pedido.lineas,
                          catalogoCompleto
                        );
                        onCopyText(`wsp-${pedido.id}`, msg, 'Reclamo para WhatsApp copiado.');
                      }}
                      className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-medium bg-white border border-red-200 text-red-700 hover:bg-red-50"
                    >
                      {copiadoId === `wsp-${pedido.id}` ? (
                        <Check className="w-3 h-3 text-emerald-600" strokeWidth={1.75} />
                      ) : (
                        <MessageCircle className="w-3 h-3" strokeWidth={1.5} />
                      )}
                      {copiadoId === `wsp-${pedido.id}` ? 'Copiado' : 'Copiar WhatsApp'}
                    </button>
                  </div>

                  {pedido.lineas
                    .filter((l) => l.requiereRevision || l.estado !== 'OK')
                    .map((l) => {
                      const prod = catalogoCompleto.find((p) => p.idGrupo === l.idGrupo);
                      const variantes = prod ? prod.variantes : l.desglose;
                      if (variantes.length < 2) return null;
                      const key = `drw_${pedido.id}_${l.idLinea}`;
                      const cur = resolucionesManual[key] ?? variantes.map(() => '0');
                      return (
                        <div key={l.idLinea} className="bg-white rounded-xl border border-red-200/70 p-3 space-y-2">
                          <div className="flex justify-between text-xs">
                            <span className="font-semibold">{l.nombreProducto}</span>
                            <span className="font-semibold text-red-700 tabular-nums">
                              Total: {l.cantidadTotal} u.
                            </span>
                          </div>
                          {l.motivoRevision && <p className="text-[11px] text-red-600">{l.motivoRevision}</p>}
                          <div className="flex flex-wrap items-end gap-2">
                            {variantes.map((v, vIdx) => (
                              <div key={v.codigo} className="flex-1 min-w-[100px]">
                                <label className="block text-[11px] text-zinc-500 mb-1">
                                  {v.etiqueta} ({v.codigo})
                                </label>
                                <input
                                  type="number"
                                  min={0}
                                  max={l.cantidadTotal}
                                  value={cur[vIdx] ?? '0'}
                                  onChange={(e) => {
                                    const nextArr = variantes.map((_, idx) =>
                                      idx === vIdx ? e.target.value : (cur[idx] ?? '0')
                                    );
                                    setResolucionesManual((prev) => ({ ...prev, [key]: nextArr }));
                                  }}
                                  className="w-full px-2.5 py-1.5 text-xs bg-white border border-zinc-200 rounded-lg tabular-nums"
                                />
                              </div>
                            ))}
                            <button
                              type="button"
                              onClick={() =>
                                onResolverLinea(
                                  pedido,
                                  l.idLinea,
                                  variantes.map((_, idx) => parseInt(cur[idx] ?? '0', 10) || 0)
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

              {esAdmin && (
                <div className="space-y-3 bg-[#FAFAFA] rounded-2xl border border-zinc-200/80 p-4 text-xs">
                  <div className="flex items-center justify-between gap-2">
                    <h3 className="text-xs font-semibold text-zinc-600 uppercase tracking-wider flex items-center gap-1.5">
                      <UserCheck className="w-3.5 h-3.5" strokeWidth={1.5} />
                      Derivación a cuentas
                    </h3>
                    <button
                      type="button"
                      onClick={() => {
                        setMostrarEnviar((prev) => !prev);
                        setRequiereConfirmSinPago(false);
                      }}
                      className="px-3 py-1 rounded-full text-[11px] font-medium bg-zinc-900 text-white hover:bg-zinc-800"
                    >
                      {mostrarEnviar ? 'Cerrar envío' : '+ Enviar a cuenta'}
                    </button>
                  </div>

                  {pedido.asignaciones.length === 0 ? (
                    <p className="text-[11px] text-zinc-500">
                      Solo en Administración · Todavía no fue enviado a Depósito ni a Ventas.
                    </p>
                  ) : (
                    <div className="space-y-1.5">
                      {pedido.asignaciones.map((a) => (
                        <div
                          key={a.id}
                          className="flex items-center justify-between gap-2 bg-white border border-zinc-200/80 rounded-xl px-3 py-2"
                        >
                          <div className="min-w-0">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <span className="font-semibold text-[#111111]">{a.usuarioNombre}</span>
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
                              <div className="text-[11px] text-zinc-500 truncate mt-0.5">{a.notaEnvio}</div>
                            )}
                          </div>
                          <button
                            type="button"
                            onClick={() => void onRetirarAsignacion(pedido.id, a.usuarioId)}
                            className="p-1 text-zinc-400 hover:text-red-600"
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
                    <div className="pt-2 border-t border-zinc-200/70 space-y-2.5">
                      <div>
                        <label className="block text-[11px] text-zinc-500 mb-1">Cuenta destino</label>
                        <select
                          value={cuentaDestinoId}
                          onChange={(e) => {
                            setCuentaDestinoId(e.target.value);
                            setRequiereConfirmSinPago(false);
                          }}
                          className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
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
                          <label className="block text-[11px] text-zinc-500 mb-1">Prioridad</label>
                          <select
                            value={prioridadEnvio}
                            onChange={(e) => setPrioridadEnvio(e.target.value as PrioridadAsignacion)}
                            className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                          >
                            <option value="NORMAL">Normal</option>
                            <option value="URGENTE">Urgente</option>
                          </select>
                        </div>
                        <div>
                          <label className="block text-[11px] text-zinc-500 mb-1">Nota breve (opc.)</label>
                          <input
                            type="text"
                            maxLength={300}
                            value={notaEnvio}
                            onChange={(e) => setNotaEnvio(e.target.value)}
                            placeholder="Ej. Sale transporte 15hs"
                            className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                          />
                        </div>
                      </div>

                      {requiereConfirmSinPago ? (
                        <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 space-y-2">
                          <p className="text-[11px] text-amber-900 font-medium">
                            Este pedido figura sin pago registrado. Confirmá si querés liberarlo igual a Depósito.
                          </p>
                          <div className="flex gap-2">
                            <button
                              type="button"
                              onClick={() => setRequiereConfirmSinPago(false)}
                              className="flex-1 py-1.5 rounded-full bg-white border border-zinc-200 text-zinc-700"
                            >
                              Cancelar
                            </button>
                            <button
                              type="button"
                              disabled={enviandoCuenta}
                              onClick={() => void handleConfirmarEnvioCuenta(true)}
                              className="flex-1 py-1.5 rounded-full bg-zinc-900 text-white font-medium"
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
                          className="w-full py-2 rounded-full bg-zinc-900 hover:bg-zinc-800 text-white font-medium disabled:opacity-40"
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
                  <h3 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">
                    Datos del cliente
                  </h3>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                    <div>
                      <label htmlFor={`drawer-cliente-${pedido.id}`} className="block text-[11px] text-zinc-500 mb-1">
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
                        className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900"
                      />
                    </div>
                    <div>
                      <label htmlFor={`drawer-whatsapp-${pedido.id}`} className="block text-[11px] text-zinc-500 mb-1">
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
                        className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl tabular-nums focus:outline-none focus:border-zinc-900"
                      />
                    </div>
                  </div>
                </div>
              )}

              {!esDeposito && (
                <div className="pt-5 border-t border-zinc-200/70 space-y-3">
                  <div className="flex items-center justify-between">
                    <h3 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">
                      Pagos y liquidación
                    </h3>
                    {pedido.estadoPago === 'Pagado' ? (
                      <StatusPill tone="green">Pagado</StatusPill>
                    ) : pedido.estadoPago === 'Pendiente' ? (
                      <StatusPill tone="amber">Pendiente</StatusPill>
                    ) : (
                      <StatusPill tone="red">Falta pago</StatusPill>
                    )}
                  </div>

                  <div className="bg-[#FAFAFA] rounded-2xl border border-zinc-200/80 p-4 space-y-1.5 text-xs tabular-nums">
                    <div className="flex justify-between text-zinc-500">
                      <span>Subtotal Neto</span>
                      <span>{formatARS(pedido.subtotalNeto)}</span>
                    </div>
                    {pedido.descuentoPorc > 0 && (
                      <div className="flex justify-between text-zinc-500">
                        <span>Descuento ({pedido.descuentoPorc}%)</span>
                        <span>-{formatARS(pedido.montoDescuento)}</span>
                      </div>
                    )}
                    <div className="flex justify-between text-zinc-500">
                      <span>Base Imponible (Total B)</span>
                      <span>{formatARS(pedido.baseImponible)}</span>
                    </div>
                    <div className="flex justify-between text-zinc-500">
                      <span>IVA 21%</span>
                      <span>{formatARS(pedido.iva21)}</span>
                    </div>
                    <div className="flex justify-between font-semibold text-sm text-[#111111] pt-1.5 border-t border-zinc-200/70">
                      <span>Total Facturado</span>
                      <span>{formatARS(pedido.totalFacturado)}</span>
                    </div>
                    <div className="flex justify-between text-emerald-700 font-medium">
                      <span>Cobrado en mano</span>
                      <span>{formatARS(pedido.montoAbonado)}</span>
                    </div>
                    <div className="flex justify-between text-zinc-700 font-medium">
                      <span>Saldo pendiente</span>
                      <span>{formatARS(pedido.saldoPendiente)}</span>
                    </div>
                  </div>

                  {esAdmin ? (
                    <div className="space-y-2.5 text-xs">
                      {pedido.medioPago === 'Cheque' && !pedido.chequeAcreditado && (
                        <button
                          type="button"
                          onClick={() => onConfirmarAcreditacion(pedido)}
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
                                void onPatch(
                                  pedido.id,
                                  {
                                    medioPago: m,
                                    montoAbonado: pedido.totalFacturado,
                                    montoCheque: 0,
                                    chequeAcreditado: true
                                  },
                                  'Marcó pago completo (Transferencia / Efectivo)'
                                );
                              } else if (m === 'Cheque') {
                                const hoy = getTodayISO();
                                void onPatch(
                                  pedido.id,
                                  {
                                    medioPago: 'Cheque',
                                    chequeAcreditado: false,
                                    montoCheque: round2(pedido.totalFacturado - pedido.montoAbonado),
                                    fechaRecepcionCheque: pedido.fechaRecepcionCheque || hoy,
                                    fechaAcreditacionCheque:
                                      pedido.fechaAcreditacionCheque ||
                                      new Date(Date.now() + 30 * 86400000).toISOString().slice(0, 10)
                                  },
                                  'Registró pago con Cheque (30 días)'
                                );
                              } else {
                                void onPatch(
                                  pedido.id,
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
                              pedido.medioPago === m
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
                            max={pedido.totalFacturado}
                            step="0.01"
                            value={draft.montoAbonado}
                            onChange={(e) => setDraft((prev) => ({ ...prev, montoAbonado: e.target.value }))}
                            onBlur={() => {
                              const val = Math.max(
                                0,
                                Math.min(pedido.totalFacturado, parseFloat(draft.montoAbonado) || 0)
                              );
                              if (Math.abs(val - pedido.montoAbonado) > 0.005) {
                                void onPatch(
                                  pedido.id,
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
                            disabled={pedido.lineas.length === 0}
                            value={draft.descuentoPorc}
                            onChange={(e) => setDraft((prev) => ({ ...prev, descuentoPorc: e.target.value }))}
                            onBlur={() => {
                              if (pedido.lineas.length === 0) {
                                onNotify('Los pedidos históricos sin detalle mantienen su total importado.', 'amber');
                                return;
                              }
                              const val = Math.max(0, Math.min(100, parseFloat(draft.descuentoPorc) || 0));
                              if (val !== pedido.descuentoPorc) {
                                void onPatch(
                                  pedido.id,
                                  { descuentoPorc: val },
                                  `Actualizó descuento a ${val}%`
                                );
                              }
                            }}
                            className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl tabular-nums disabled:bg-zinc-50 disabled:text-zinc-400"
                          />
                        </div>
                      </div>

                      {pedido.medioPago === 'Cheque' && !pedido.chequeAcreditado && (
                        <div className="grid grid-cols-2 gap-3 pt-1">
                          <div>
                            <label className="block text-[11px] text-zinc-500 mb-1">Fecha acreditación cheque</label>
                            <input
                              type="date"
                              value={draft.fechaAcreditacionCheque}
                              onChange={(e) =>
                                setDraft((prev) => ({ ...prev, fechaAcreditacionCheque: e.target.value }))
                              }
                              onBlur={() => {
                                const f = draft.fechaAcreditacionCheque;
                                if (f && f !== pedido.fechaAcreditacionCheque) {
                                  void onPatch(
                                    pedido.id,
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
                              value={draft.detalleCheque}
                              onChange={(e) => setDraft((prev) => ({ ...prev, detalleCheque: e.target.value }))}
                              onBlur={() => {
                                const d = draft.detalleCheque.trim();
                                if (d !== pedido.detalleCheque) {
                                  void onPatch(
                                    pedido.id,
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
              )}

              <div className="pt-5 border-t border-zinc-200/70 space-y-3">
                <div className="flex items-center justify-between gap-2">
                  <h3 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">Armado y envío</h3>
                  {draft.transporte.trim() && draft.nroSeguimiento.trim() && (
                    <button
                      type="button"
                      onClick={() => {
                        const msg = generarMensajeDespachoWhatsApp({
                          cliente: pedido.cliente,
                          numeroPedido: pedido.numeroPedido,
                          transporte: draft.transporte.trim(),
                          nroSeguimiento: draft.nroSeguimiento.trim()
                        });
                        if (msg) onCopyText(`envio-${pedido.id}`, msg, 'Aviso de envío copiado para WhatsApp');
                      }}
                      className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-medium bg-zinc-100 hover:bg-zinc-200/80 text-zinc-700 transition-colors"
                    >
                      {copiadoId === `envio-${pedido.id}` ? (
                        <Check className="w-3 h-3 text-emerald-600" strokeWidth={1.75} />
                      ) : (
                        <Send className="w-3 h-3" strokeWidth={1.5} />
                      )}
                      {copiadoId === `envio-${pedido.id}` ? 'Aviso copiado' : 'Copiar aviso de envío'}
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-2 gap-3 text-xs">
                  <div>
                    <label className="block text-[11px] text-zinc-500 mb-1">Estado de armado</label>
                    {pedido.faltantesVariante > 0 ? (
                      <div className="px-3 py-2 rounded-xl bg-red-50 text-red-700 border border-red-200/70 font-medium">
                        Falta talle o color
                      </div>
                    ) : (
                      <select
                        value={pedido.estadoArmado}
                        onChange={(e) => {
                          const nuevoEstado = e.target.value as EstadoArmado;
                          void onPatch(pedido.id, { estadoArmado: nuevoEstado }, `Cambió armado a ${nuevoEstado}`);
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
                      value={pedido.estadoEnvio}
                      onChange={(e) => {
                        const nuevoEnvio = e.target.value as EstadoEnvio;
                        void onPatch(pedido.id, { estadoEnvio: nuevoEnvio }, `Cambió envío a ${nuevoEnvio}`);
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
                      value={draft.transporte}
                      onChange={(e) => setDraft((prev) => ({ ...prev, transporte: e.target.value }))}
                      onBlur={() => {
                        const val = draft.transporte.trim();
                        if (val !== pedido.transporte) {
                          void onPatch(pedido.id, { transporte: val }, `Transporte: ${val || 'vacío'}`);
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
                      value={draft.nroSeguimiento}
                      onChange={(e) => setDraft((prev) => ({ ...prev, nroSeguimiento: e.target.value }))}
                      onBlur={() => {
                        const val = draft.nroSeguimiento.trim();
                        if (val !== pedido.nroSeguimiento) {
                          void onPatch(
                            pedido.id,
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
                      value={draft.observaciones}
                      onChange={(e) => setDraft((prev) => ({ ...prev, observaciones: e.target.value }))}
                      onBlur={() => {
                        const val = draft.observaciones.trim();
                        if (val !== pedido.observaciones) {
                          void onPatch(pedido.id, { observaciones: val }, 'Actualizó observaciones');
                        }
                      }}
                      className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                    />
                  </div>
                </div>

                {esDeposito && (
                  <div className="pt-2">
                    {!mostrarIncidenciaForm ? (
                      <button
                        type="button"
                        onClick={() => setMostrarIncidenciaForm(true)}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-amber-50 border border-amber-200 text-amber-800 hover:bg-amber-100/80"
                      >
                        <AlertTriangle className="w-3.5 h-3.5" strokeWidth={1.5} />
                        Reportar faltante / problema
                      </button>
                    ) : (
                      <div className="bg-amber-50/70 border border-amber-200 rounded-xl p-3 space-y-2 text-xs">
                        <label className="block text-[11px] font-medium text-amber-900">
                          Detalle para Administración
                        </label>
                        <textarea
                          rows={2}
                          value={incidenciaTexto}
                          onChange={(e) => setIncidenciaTexto(e.target.value)}
                          placeholder="Ej. Faltan 2 Chalecos Shark T1 en estantería..."
                          className="w-full px-3 py-2 bg-white border border-amber-200 rounded-lg"
                        />
                        <div className="flex justify-end gap-2">
                          <button
                            type="button"
                            onClick={() => setMostrarIncidenciaForm(false)}
                            className="px-3 py-1.5 rounded-full bg-white border border-zinc-200 text-zinc-600"
                          >
                            Cancelar
                          </button>
                          <button
                            type="button"
                            disabled={enviandoIncidencia}
                            onClick={() => void handleEnviarIncidencia()}
                            className="px-3.5 py-1.5 rounded-full bg-zinc-900 text-white font-medium"
                          >
                            {enviandoIncidencia ? 'Enviando...' : 'Enviar aviso'}
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {pedido.lineas.length > 0 && (
                <div className="pt-5 border-t border-zinc-200/70 space-y-2.5">
                  <h3 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">
                    Detalle de artículos ({pedido.lineas.length} · {totalUnidades} u.)
                  </h3>
                  <div className="divide-y divide-zinc-200/60 border border-zinc-200/80 rounded-xl">
                    {pedido.lineas.map((l) => (
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
                          {esDeposito
                            ? `${l.cantidadTotal} u.`
                            : formatARS(round2(l.precioUnitarioNeto * l.cantidadTotal))}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              <div className="pt-5 border-t border-zinc-200/70 space-y-2.5">
                <h3 className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">
                  Historial de cambios
                </h3>
                <ul className="space-y-1.5 text-[11px] text-zinc-500">
                  {pedido.historial.map((h, i) => (
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

              {esAdmin && (
                <div className="pt-5 border-t border-zinc-200/70 flex justify-end">
                  {confirmEliminar ? (
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => setConfirmEliminar(false)}
                        className="px-3 py-1.5 rounded-full text-xs border border-zinc-200 text-zinc-600"
                      >
                        Cancelar
                      </button>
                      <button
                        type="button"
                        onClick={() => onEliminarIndividual(pedido)}
                        className="px-3.5 py-1.5 rounded-full text-xs font-medium bg-red-600 text-white"
                      >
                        Confirmar eliminación
                      </button>
                    </div>
                  ) : (
                    <button
                      type="button"
                      onClick={() => setConfirmEliminar(true)}
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
  );
}
