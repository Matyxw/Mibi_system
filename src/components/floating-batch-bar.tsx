'use client';

import React, { useState } from 'react';
import { CreditCard, Download, FileSpreadsheet, Layers, Lock, Trash2, Truck, X } from 'lucide-react';
import { VENDEDORES_MIBI } from '@/lib/catalogo-mibi';
import { formatARS } from '@/lib/finanzas';
import type { BatchPanelActivo, Pedido } from '@/lib/types';

interface FloatingBatchBarProps {
  selectedCount: number;
  totalSeleccionado: number;
  pedidosSeleccionados: ReadonlyArray<Pedido>;
  esAdmin: boolean;
  batchPanel: BatchPanelActivo;
  setBatchPanel: React.Dispatch<React.SetStateAction<BatchPanelActivo>>;
  onLimpiarSeleccion: () => void;
  onBatchArmado: (destino: 'En armado' | 'Listo') => void;
  onBatchEnvio: (destino: 'Despachado' | 'Entregado', transporte: string) => void;
  onBatchPago: (accion: 'Pagado' | 'Falta pago') => void;
  onBatchVendedor: (vendedor: string) => void;
  onExportTableroSeleccionados: () => void;
  onExportPickingConsolidado: () => void;
  onConfirmarEliminarLote: () => void;
}

export function FloatingBatchBar({
  selectedCount,
  totalSeleccionado,
  pedidosSeleccionados,
  esAdmin,
  batchPanel,
  setBatchPanel,
  onLimpiarSeleccion,
  onBatchArmado,
  onBatchEnvio,
  onBatchPago,
  onBatchVendedor,
  onExportTableroSeleccionados,
  onExportPickingConsolidado,
  onConfirmarEliminarLote
}: FloatingBatchBarProps): React.JSX.Element {
  const [batchTransporte, setBatchTransporte] = useState<string>('');
  const [batchVendedor, setBatchVendedor] = useState<string>('Juancho');

  return (
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
                  onClick={() => onBatchArmado('En armado')}
                  className="px-3 py-2 rounded-xl text-xs font-medium bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 text-zinc-800 transition-colors"
                >
                  En armado
                </button>
                <button
                  type="button"
                  onClick={() => onBatchArmado('Listo')}
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
                  onClick={() => {
                    onBatchEnvio('Despachado', batchTransporte);
                    setBatchTransporte('');
                  }}
                  className="px-3 py-2 rounded-xl text-xs font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors"
                >
                  Marcar Despachado
                </button>
                <button
                  type="button"
                  onClick={() => {
                    onBatchEnvio('Entregado', batchTransporte);
                    setBatchTransporte('');
                  }}
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
                onClick={() => onBatchPago('Pagado')}
                className="w-full px-3 py-2 rounded-xl text-xs font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors"
              >
                Marcar Pagado (Cobrado en mano)
              </button>
              <button
                type="button"
                onClick={() => onBatchPago('Falta pago')}
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
                  onClick={() => onBatchVendedor(batchVendedor)}
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
              onClick={onExportTableroSeleccionados}
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
              onClick={onExportPickingConsolidado}
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
                onClick={onConfirmarEliminarLote}
                className="flex-1 py-2 rounded-full text-xs font-medium bg-red-600 hover:bg-red-700 text-white"
              >
                Confirmar
              </button>
            </div>
          </div>
        )}

        <div className="bg-zinc-900 text-white rounded-full shadow-lg px-3.5 py-2 flex items-center gap-2 sm:gap-3 border border-zinc-800">
          <div className="flex items-center gap-2 pl-1 whitespace-nowrap">
            <span className="inline-flex items-center justify-center min-w-[20px] h-5 px-1.5 rounded-full bg-[#F4B6CE] text-[#111111] text-[11px] font-bold tabular-nums">
              {selectedCount}
            </span>
            <span className="text-xs font-medium hidden sm:inline">
              {selectedCount === 1 ? 'seleccionado' : 'seleccionados'}
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
                <span className="hidden sm:inline">Eliminar ({selectedCount})</span>
              </button>
            )}
          </div>

          <button
            type="button"
            onClick={onLimpiarSeleccion}
            className="p-1.5 rounded-full text-zinc-400 hover:text-white hover:bg-zinc-800 transition-colors"
            aria-label="Limpiar selección (Escape)"
            title="Deseleccionar (Esc)"
          >
            <X className="w-3.5 h-3.5" strokeWidth={1.75} />
          </button>
        </div>
      </div>
    </div>
  );
}
