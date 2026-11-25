'use client';

import React, { useEffect, useRef } from 'react';
import { ChevronRight } from 'lucide-react';
import {
  esDespachoAtrasado,
  formatARS,
  formatARSCompacto,
  formatShortDate,
  getDiasParaAcreditacionCheque
} from '@/lib/finanzas';
import { isDepositoRole, type Pedido } from '@/lib/types';
import { StatusPill } from './status-pill';

interface PedidosTableProps {
  pedidosFiltrados: ReadonlyArray<Pedido>;
  selectedIds: ReadonlySet<string>;
  todosVisiblesSeleccionados: boolean;
  seleccionParcialVisible: boolean;
  esAdmin: boolean;
  esDeposito: boolean;
  onToggleSelectAll: () => void;
  onToggleSelectRow: (id: string, shiftKey: boolean) => void;
  onOpenDrawer: (id: string) => void;
  onConfirmarAcreditacion: (pedido: Pedido) => void;
  onExportarHojaDeposito: (pedido: Pedido) => void;
}

export function PedidosTable({
  pedidosFiltrados,
  selectedIds,
  todosVisiblesSeleccionados,
  seleccionParcialVisible,
  esAdmin,
  esDeposito,
  onToggleSelectAll,
  onToggleSelectRow,
  onOpenDrawer,
  onConfirmarAcreditacion,
  onExportarHojaDeposito
}: PedidosTableProps): React.JSX.Element {
  const headerCheckboxRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (headerCheckboxRef.current) {
      headerCheckboxRef.current.indeterminate = seleccionParcialVisible;
    }
  }, [seleccionParcialVisible]);

  return (
    <div className="bg-white dark:bg-[#141417] rounded-2xl border border-[#EAEAEA] dark:border-[#26262B] shadow-[0_1px_2px_rgba(0,0,0,0.03)] overflow-hidden">
      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="border-b border-zinc-200/70 dark:border-[#26262B] text-[11px] font-medium text-zinc-400 dark:text-zinc-500">
              <th className="py-3.5 pl-5 pr-2 w-10">
                <input
                  ref={headerCheckboxRef}
                  type="checkbox"
                  checked={todosVisiblesSeleccionados}
                  onChange={onToggleSelectAll}
                  disabled={pedidosFiltrados.length === 0}
                  aria-label="Seleccionar todos los pedidos visibles"
                  className="w-4 h-4 rounded border-zinc-300 dark:border-zinc-600 text-zinc-900 accent-zinc-900 dark:accent-[#F4B6CE] cursor-pointer align-middle disabled:opacity-30"
                />
              </th>
              <th className="py-3.5 pl-2 pr-3">N.º y Fecha</th>
              <th className="py-3.5 px-3">{esDeposito ? 'Cliente y Nota' : 'Cliente'}</th>
              {esDeposito ? (
                <>
                  <th className="py-3.5 px-3">Artículos / Unidades</th>
                  <th className="py-3.5 px-3">Despacho previsto</th>
                  <th className="py-3.5 px-3">Armado y Envío</th>
                </>
              ) : (
                <>
                  <th className="py-3.5 px-3 text-right">Total</th>
                  <th className="py-3.5 px-3">Pago</th>
                  <th className="py-3.5 px-3">Armado</th>
                </>
              )}
              <th className="py-3.5 pl-3 pr-6 text-right">Acción</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-zinc-200/60 dark:divide-[#26262B]/80 text-sm">
            {pedidosFiltrados.length === 0 ? (
              <tr>
                <td colSpan={7} className="py-14 text-center text-sm text-zinc-400 dark:text-zinc-500">
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
                const totalUnidades = p.lineas.reduce((acc, l) => acc + l.cantidadTotal, 0);

                const asigsDeposito = (p.asignaciones ?? []).filter((a) => isDepositoRole(a.usuarioRol));
                const esUrgente = asigsDeposito.some((a) => a.prioridad === 'URGENTE');
                const requiereRelectura = asigsDeposito.some((a) => a.requiereRelectura);
                const vistoDeposito = asigsDeposito.some((a) => a.vistoAt !== null);
                const notaDeposito = Array.from(
                  new Set(asigsDeposito.map((a) => a.notaEnvio.trim()).filter(Boolean))
                ).join(' · ');

                return (
                  <tr
                    key={p.id}
                    onClick={() => onOpenDrawer(p.id)}
                    className={`cursor-pointer transition-colors ${
                      seleccionado
                        ? 'bg-pink-50/45 hover:bg-pink-50/70 dark:bg-[#F4B6CE]/[0.08] dark:hover:bg-[#F4B6CE]/[0.12]'
                        : 'hover:bg-zinc-50/80 dark:hover:bg-[#1B1B1F]'
                    }`}
                  >
                    <td
                      className="py-3.5 pl-5 pr-2 w-10"
                      onClick={(e) => {
                        e.stopPropagation();
                        onToggleSelectRow(p.id, e.shiftKey);
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={seleccionado}
                        onChange={() => {}}
                        aria-label={`Seleccionar pedido #${p.numeroPedido}`}
                        className="w-4 h-4 rounded border-zinc-300 dark:border-zinc-600 text-zinc-900 accent-zinc-900 dark:accent-[#F4B6CE] cursor-pointer align-middle pointer-events-none"
                      />
                    </td>

                    <td className="py-3.5 pl-2 pr-3 whitespace-nowrap tabular-nums">
                      <div className="font-semibold text-[#111111] dark:text-[#F4F4F5]">#{p.numeroPedido}</div>
                      <div className="text-xs text-zinc-400 dark:text-zinc-500 mt-0.5 flex items-center gap-1.5">
                        <span>{formatShortDate(p.fechaPedido)}</span>
                        {atrasado && (
                          <span
                            className="w-1.5 h-1.5 rounded-full bg-amber-500 dark:bg-amber-400"
                            title={`Despacho previsto: ${formatShortDate(p.despachoPrevisto)}`}
                          />
                        )}
                      </div>
                    </td>

                    <td className="py-3.5 px-3 min-w-[200px]">
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <span className="font-medium text-[#111111] dark:text-[#F4F4F5]">{p.cliente}</span>
                        {esDeposito && esUrgente && <StatusPill tone="red">Urgente</StatusPill>}
                        {esDeposito && requiereRelectura && <StatusPill tone="amber">Actualizado</StatusPill>}
                      </div>
                      <div className="text-xs text-zinc-400 dark:text-zinc-500 mt-0.5">
                        {esDeposito
                          ? notaDeposito || p.vendedor
                          : `${p.vendedor}${p.detalleResumen && p.detalleResumen !== '—' ? ` · ${p.detalleResumen}` : ''}`}
                      </div>
                    </td>

                    {esDeposito ? (
                      <>
                        <td className="py-3.5 px-3 whitespace-nowrap tabular-nums">
                          <div className="font-semibold text-[#111111] dark:text-[#F4F4F5]">{totalUnidades} u.</div>
                          <div className="text-xs text-zinc-400 dark:text-zinc-500 mt-0.5">{p.lineas.length} artículos</div>
                        </td>

                        <td className="py-3.5 px-3 whitespace-nowrap tabular-nums">
                          <div className="text-xs font-medium text-[#111111] dark:text-[#F4F4F5]">
                            {p.despachoPrevisto ? formatShortDate(p.despachoPrevisto) : '—'}
                          </div>
                          {atrasado && <div className="text-[11px] text-amber-700 dark:text-amber-400 mt-0.5">Atrasado</div>}
                        </td>
                      </>
                    ) : (
                      <>
                        <td className="py-3.5 px-3 text-right whitespace-nowrap tabular-nums">
                          <div className="font-semibold text-[#111111] dark:text-[#F4F4F5]">{formatARS(p.totalFacturado)}</div>
                          <div className="text-xs text-zinc-400 dark:text-zinc-500 mt-0.5">
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
                                onClick={() => onConfirmarAcreditacion(p)}
                                className="text-[11px] font-medium text-emerald-700 hover:text-emerald-800 dark:text-emerald-400 dark:hover:text-emerald-300 underline underline-offset-2"
                              >
                                Confirmar plata acreditada
                              </button>
                            )}
                          </div>
                        </td>
                      </>
                    )}

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
                          <span className="text-[11px] text-zinc-400 dark:text-zinc-500 pl-0.5">
                            {p.estadoEnvio}
                            {p.transporte ? ` · ${p.transporte}` : ''}
                          </span>
                        ) : atrasado && !esDeposito ? (
                          <span className="text-[11px] text-amber-700 dark:text-amber-400 pl-0.5">Despacho atrasado</span>
                        ) : null}

                        {!esDeposito && (
                          <span className="text-[11px] pl-0.5">
                            {p.incidenciaDeposito ? (
                              <span className="text-amber-700 dark:text-amber-400 font-medium">Aviso de depósito</span>
                            ) : asigsDeposito.length > 0 ? (
                              <span className={esUrgente ? 'text-red-600 dark:text-red-400 font-medium' : 'text-zinc-400 dark:text-zinc-500'}>
                                {esUrgente
                                  ? 'Urgente · Depósito'
                                  : requiereRelectura
                                    ? 'En Depósito · Actualizado'
                                    : vistoDeposito
                                      ? 'En Depósito · Visto'
                                      : 'En Depósito · Sin leer'}
                              </span>
                            ) : (
                              <span className="text-zinc-400 dark:text-zinc-500">
                                {esAdmin ? 'Solo en Administración' : 'En seguimiento'}
                              </span>
                            )}
                          </span>
                        )}
                      </div>
                    </td>

                    <td
                      className="py-3.5 pl-3 pr-6 text-right whitespace-nowrap"
                      onClick={(e) => e.stopPropagation()}
                    >
                      <div className="inline-flex items-center gap-1.5">
                        {faltaDefinir && !esDeposito ? (
                          <button
                            type="button"
                            onClick={() => onOpenDrawer(p.id)}
                            className="px-3 py-1.5 rounded-full text-xs font-medium bg-red-50 text-red-700 border border-red-200/70 hover:bg-red-100/70 dark:bg-red-500/[0.12] dark:text-red-300 dark:border-red-500/25 dark:hover:bg-red-500/20 transition-colors"
                          >
                            Completar
                          </button>
                        ) : (
                          <button
                            type="button"
                            onClick={() => onExportarHojaDeposito(p)}
                            className="px-3 py-1.5 rounded-full text-xs font-medium bg-white text-zinc-700 border border-zinc-200 hover:bg-zinc-50 dark:bg-[#1B1B1F] dark:text-zinc-200 dark:border-[#2E2E35] dark:hover:bg-[#24242B] transition-colors"
                          >
                            Hoja depósito
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={() => onOpenDrawer(p.id)}
                          className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-800 hover:bg-zinc-100 dark:hover:text-[#F4F4F5] dark:hover:bg-[#24242B] transition-colors"
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
  );
}
