'use client';

import React, { useMemo, useRef, useState } from 'react';
import { AlertTriangle, ArrowLeft, Check, FileSpreadsheet, Lock, MessageCircle, X } from 'lucide-react';
import { CATALOGO_BY_GRUPO, CATALOGO_MIBI, VENDEDORES_MIBI } from '@/lib/catalogo-mibi';
import {
  aplicarResolucionLinea,
  crearLineasEjemploLista006,
  normalizarTexto,
  parsearExcelLista006,
  parsearMontoArgentino,
  resolverLinea
} from '@/lib/excel-parser';
import { addDaysISO, formatARS, formatARSCompacto, fromCents, getTodayISO, round2, toCents } from '@/lib/finanzas';
import { generarMensajeWhatsApp } from '@/lib/whatsapp';
import type {
  AdvertenciaExcel,
  LineaPedido,
  MedioPago,
  ProductoCatalogo,
  ToneStatus
} from '@/lib/types';

export interface FormNuevoPedidoState {
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

export const crearFormInicial = (): FormNuevoPedidoState => {
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

interface NuevoPedidoViewProps {
  esAdmin: boolean;
  proximoNumeroPedido: string;
  form: FormNuevoPedidoState;
  setForm: React.Dispatch<React.SetStateAction<FormNuevoPedidoState>>;
  lineasBorrador: LineaPedido[];
  setLineasBorrador: React.Dispatch<React.SetStateAction<LineaPedido[]>>;
  cantidadesManuales: Record<string, number>;
  setCantidadesManuales: React.Dispatch<React.SetStateAction<Record<string, number>>>;
  advertenciasExcel: AdvertenciaExcel[];
  setAdvertenciasExcel: React.Dispatch<React.SetStateAction<AdvertenciaExcel[]>>;
  nombreArchivoCargado: string;
  setNombreArchivoCargado: React.Dispatch<React.SetStateAction<string>>;
  copiadoId: string | null;
  guardandoPedido: boolean;
  onVolver: () => void;
  onCopyText: (key: string, text: string, msgOk: string) => void;
  onNotify: (texto: string, tone?: ToneStatus) => void;
  onGuardar: (payload: {
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
  }) => Promise<void>;
}

export function NuevoPedidoView({
  esAdmin,
  proximoNumeroPedido,
  form,
  setForm,
  lineasBorrador,
  setLineasBorrador,
  cantidadesManuales,
  setCantidadesManuales,
  advertenciasExcel,
  setAdvertenciasExcel,
  nombreArchivoCargado,
  setNombreArchivoCargado,
  copiadoId,
  guardandoPedido,
  onVolver,
  onCopyText,
  onNotify,
  onGuardar
}: NuevoPedidoViewProps): React.JSX.Element {
  const [modoCargaManual, setModoCargaManual] = useState<boolean>(false);
  const [busquedaCatalogo, setBusquedaCatalogo] = useState<string>('');
  const [arrastrandoArchivo, setArrastrandoArchivo] = useState<boolean>(false);
  const [resolucionesManual, setResolucionesManual] = useState<Record<string, { q0: string; q1: string }>>({});
  const fileInputListaRef = useRef<HTMLInputElement | null>(null);

  const totalesBorrador = useMemo(() => {
    let subtotalCents = 0;
    for (const l of lineasBorrador) {
      subtotalCents += toCents(l.precioUnitarioNeto) * l.cantidadTotal;
    }
    const desc = Math.max(0, Math.min(100, Number.isFinite(form.descuentoPorc) ? round2(form.descuentoPorc) : 0));
    const montoDescuentoCents = Math.floor((subtotalCents * desc + 50) / 100);
    const baseCents = subtotalCents - montoDescuentoCents;
    const ivaCents = Math.floor((baseCents * 21 + 50) / 100);
    const totalCents = baseCents + ivaCents;
    const faltantes = lineasBorrador.filter((l) => l.requiereRevision || l.estado !== 'OK').length;
    const totalUnidades = lineasBorrador.reduce((acc, l) => acc + l.cantidadTotal, 0);

    return {
      subtotalNeto: fromCents(subtotalCents),
      montoDescuento: fromCents(montoDescuentoCents),
      baseImponible: fromCents(baseCents),
      iva21: fromCents(ivaCents),
      totalFacturado: fromCents(totalCents),
      faltantes,
      totalUnidades
    };
  }, [lineasBorrador, form.descuentoPorc]);

  const catalogoFiltradoManual = useMemo(
    () =>
      CATALOGO_MIBI.filter((prod) => {
        const q = normalizarTexto(busquedaCatalogo);
        if (!q) return true;
        return (
          normalizarTexto(prod.nombre).includes(q) ||
          normalizarTexto(prod.rubro).includes(q) ||
          prod.variantes.some((v) => v.codigo.includes(q))
        );
      }),
    [busquedaCatalogo]
  );

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
    try {
      const buffer = await file.arrayBuffer();
      const res = parsearExcelLista006(buffer, file.name);
      if (!res.ok) {
        onNotify(res.errorFatal ?? 'Archivo inválido.', 'red');
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
        onNotify(
          `Se detectaron ${res.advertencias.length} fila(s) con error para revisar antes de guardar.`,
          'red'
        );
      } else {
        onNotify(`${res.lineas.length} artículos leídos de ${file.name}.`, 'green');
      }
    } catch {
      onNotify('No se pudo leer el archivo seleccionado.', 'red');
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
    onNotify('Planilla de ejemplo Lista N° 006 cargada ($3.560.673,05 total).', 'green');
  };

  const handleResolverEnBorrador = (idLinea: string, q0: number, q1: number): void => {
    const lineaObj = lineasBorrador.find((l) => l.idLinea === idLinea);
    if (!lineaObj) return;
    const resuelta = aplicarResolucionLinea(lineaObj, q0, q1);
    if (!resuelta) {
      onNotify(`La suma debe ser exactamente ${lineaObj.cantidadTotal} u.`, 'red');
      return;
    }
    setLineasBorrador((prev) => prev.map((l) => (l.idLinea === idLinea ? resuelta : l)));
    onNotify(`${lineaObj.nombreProducto} confirmado.`, 'green');
  };

  const handleSubmit = async (): Promise<void> => {
    if (!form.cliente.trim()) {
      onNotify('Ingresá el nombre del cliente o negocio.', 'amber');
      return;
    }
    if (lineasBorrador.length === 0) {
      onNotify('Cargá al menos un artículo desde Excel o en carga manual.', 'amber');
      return;
    }
    if (advertenciasExcel.length > 0) {
      onNotify('Confirmá o descartá las alertas de Excel antes de guardar.', 'red');
      return;
    }

    const totalEst = totalesBorrador.totalFacturado;
    const medioPagoEfectivo: MedioPago = !esAdmin ? 'Sin pagar' : form.pagoTipo;

    let montoAbonado = 0;
    let montoCheque = 0;

    if (esAdmin && medioPagoEfectivo === 'Transferencia / Efectivo') {
      const rawStr = form.montoAbonadoInput.trim();
      if (!rawStr) {
        montoAbonado = totalEst;
      } else {
        const val = parsearMontoArgentino(rawStr);
        montoAbonado = val > 0 ? Math.min(totalEst, round2(val)) : 0;
      }
    } else if (esAdmin && medioPagoEfectivo === 'Cheque') {
      const rawAbon = form.montoAbonadoInput.trim();
      const valAbon = rawAbon ? parsearMontoArgentino(rawAbon) : 0;
      montoAbonado = valAbon > 0 ? Math.min(totalEst, round2(valAbon)) : 0;
      const rawCheq = form.montoChequeInput.trim();
      const valCheq = rawCheq ? parsearMontoArgentino(rawCheq) : 0;
      montoCheque =
        valCheq > 0 ? Math.min(round2(totalEst - montoAbonado), round2(valCheq)) : round2(totalEst - montoAbonado);
    }

    await onGuardar({
      cliente: form.cliente.trim(),
      vendedor: form.vendedor,
      whatsapp: form.whatsapp.trim(),
      fechaPedido: form.fechaPedido || getTodayISO(),
      despachoPrevisto: form.despachoPrevisto || addDaysISO(getTodayISO(), 5),
      observaciones: form.observaciones.trim(),
      descuentoPorc: esAdmin ? form.descuentoPorc : 0,
      medioPago: medioPagoEfectivo,
      montoAbonado,
      montoCheque,
      fechaRecepcionCheque: medioPagoEfectivo === 'Cheque' ? form.fechaRecepcionCheque : '',
      fechaAcreditacionCheque: medioPagoEfectivo === 'Cheque' ? form.fechaAcreditacionCheque : '',
      detalleCheque: medioPagoEfectivo === 'Cheque' ? form.detalleCheque.trim() : '',
      lineas: lineasBorrador
    });
  };

  return (
    <div className="max-w-2xl mx-auto space-y-6">
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

      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={onVolver}
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
              disabled={!esAdmin}
              value={esAdmin ? form.descuentoPorc : 0}
              onChange={(e) =>
                setForm((prev) => ({
                  ...prev,
                  descuentoPorc: Math.max(0, Math.min(100, Number(e.target.value) || 0))
                }))
              }
              className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 tabular-nums disabled:bg-zinc-50 disabled:text-zinc-400"
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
            <p className="mt-1 text-xs text-zinc-400">Soporta .xlsx, .xls y .csv · Lectura exacta inmediata</p>
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
                onCopyText('borrador-wsp', msg, 'Mensaje de consulta copiado para WhatsApp.');
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
                const v0 = prod.variantes[0]!;
                const v1 = prod.variantes[1]!;
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
            onClick={() => void handleSubmit()}
            className="w-full py-3 rounded-full text-sm font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors disabled:opacity-50"
          >
            {guardandoPedido ? 'Guardando...' : 'Guardar pedido'}
          </button>
        </section>
      )}
    </div>
  );
}
