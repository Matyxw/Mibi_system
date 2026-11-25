'use client';

import React, { useId, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  ArrowLeft,
  Check,
  Copy,
  Edit3,
  FileSpreadsheet,
  Lock,
  MessageCircle,
  Plus,
  Trash2,
  X
} from 'lucide-react';
import { crearProductoCatalogoAction } from '@/app/actions/catalogo-actions';
import { VENDEDORES_MIBI } from '@/lib/catalogo-mibi';
import {
  aplicarResolucionLinea,
  crearLineasEjemploLista006,
  normalizarTexto,
  parsearExcelLista006,
  parsearMontoArgentino,
  resolverLinea
} from '@/lib/excel-parser';
import { addDaysISO, formatARS, formatARSCompacto, fromCents, getTodayISO, round2, toCents } from '@/lib/finanzas';
import { buildWhatsAppUrl, generarMensajeWhatsApp } from '@/lib/whatsapp';
import type {
  AdvertenciaExcel,
  ClienteConocido,
  DesgloseVariante,
  LineaPedido,
  MedioPago,
  ProductoCatalogo,
  ProductoCustomRecord,
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

interface CustomVarianteDraft {
  codigo: string;
  etiqueta: string;
  cantidad: string;
}

interface NuevoPedidoViewProps {
  esAdmin: boolean;
  proximoNumeroPedido: string;
  catalogoCompleto: ReadonlyArray<ProductoCatalogo>;
  clientesConocidos: ReadonlyArray<ClienteConocido>;
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
  onProductoCustomCreado: (rec: ProductoCustomRecord) => void;
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
  catalogoCompleto,
  clientesConocidos,
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
  onProductoCustomCreado,
  onGuardar
}: NuevoPedidoViewProps): React.JSX.Element {
  const clientesListId = useId();
  const [modoCargaManual, setModoCargaManual] = useState<boolean>(false);
  const [busquedaCatalogo, setBusquedaCatalogo] = useState<string>('');
  const [soloCargados, setSoloCargados] = useState<boolean>(false);
  const [arrastrandoArchivo, setArrastrandoArchivo] = useState<boolean>(false);
  const [resolucionesManual, setResolucionesManual] = useState<Record<string, string[]>>({});
  const [editandoLineaId, setEditandoLineaId] = useState<string | null>(null);
  const [confirmDescartar, setConfirmDescartar] = useState<boolean>(false);
  const [confirmVaciarArticulos, setConfirmVaciarArticulos] = useState<boolean>(false);
  const fileInputListaRef = useRef<HTMLInputElement | null>(null);

  // Estado formulario "+ Producto fuera de planilla"
  const [mostrarFueraPlanilla, setMostrarFueraPlanilla] = useState<boolean>(false);
  const [fpNombre, setFpNombre] = useState<string>('');
  const [fpRubroSel, setFpRubroSel] = useState<string>('PILUSOS');
  const [fpRubroCustom, setFpRubroCustom] = useState<string>('');
  const [fpPrecioNeto, setFpPrecioNeto] = useState<string>('');
  const [fpTipoVar, setFpTipoVar] = useState<'unico' | 'multiple'>('unico');
  const [fpCodigoUnico, setFpCodigoUnico] = useState<string>('');
  const [fpCantidadTotal, setFpCantidadTotal] = useState<string>('1');
  const [fpVariantes, setFpVariantes] = useState<CustomVarianteDraft[]>([
    { codigo: '', etiqueta: 'T1 (1-3 años)', cantidad: '' },
    { codigo: '', etiqueta: 'T2 (3-6 años)', cantidad: '' }
  ]);
  const [fpGuardarEnCatalogo, setFpGuardarEnCatalogo] = useState<boolean>(true);
  const [fpGuardando, setFpGuardando] = useState<boolean>(false);

  const rubrosDisponibles = useMemo(
    () => Array.from(new Set(['PILUSOS', ...catalogoCompleto.map((p) => p.rubro), 'ACCESORIOS'])),
    [catalogoCompleto]
  );

  const clienteCoincidente = useMemo<ClienteConocido | null>(() => {
    const q = form.cliente.trim().toLowerCase();
    if (!q) return null;
    return clientesConocidos.find((c) => c.cliente.toLowerCase() === q) ?? null;
  }, [clientesConocidos, form.cliente]);

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

  const cargadosManualCount = useMemo(
    () => Object.values(cantidadesManuales).filter((q) => q > 0).length,
    [cantidadesManuales]
  );

  const catalogoFiltradoManual = useMemo(
    () =>
      catalogoCompleto.filter((prod) => {
        if (soloCargados && (cantidadesManuales[prod.idGrupo] ?? 0) <= 0) return false;
        const q = normalizarTexto(busquedaCatalogo);
        if (!q) return true;
        return (
          normalizarTexto(prod.nombre).includes(q) ||
          normalizarTexto(prod.rubro).includes(q) ||
          prod.variantes.some((v) => normalizarTexto(v.codigo).includes(q))
        );
      }),
    [catalogoCompleto, busquedaCatalogo, soloCargados, cantidadesManuales]
  );

  const sincronizarManualConBorrador = (nextMap: Record<string, number>): void => {
    setLineasBorrador((prev) => {
      const prevByGrupo = new Map<string, LineaPedido>();
      for (const l of prev) prevByGrupo.set(l.idGrupo, l);
      const catIds = new Set(catalogoCompleto.map((c) => c.idGrupo));

      const resultado: LineaPedido[] = [];
      for (const prod of catalogoCompleto) {
        const qty = nextMap[prod.idGrupo] ?? 0;
        if (qty <= 0) continue;
        const existente = prevByGrupo.get(prod.idGrupo);
        if (existente && existente.cantidadTotal === qty) {
          resultado.push(existente);
        } else {
          resultado.push(resolverLinea(prod, qty, existente?.notaCliente ?? ''));
        }
      }
      for (const l of prev) {
        if (!catIds.has(l.idGrupo) && (nextMap[l.idGrupo] ?? l.cantidadTotal) > 0) {
          resultado.push(l);
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

  const handleDescartarBorrador = (): void => {
    setForm(crearFormInicial());
    setLineasBorrador([]);
    setCantidadesManuales({});
    setAdvertenciasExcel([]);
    setNombreArchivoCargado('');
    setSoloCargados(false);
    setConfirmDescartar(false);
    setConfirmVaciarArticulos(false);
    onNotify('Borrador descartado.', 'neutral');
  };

  const handleArchivoLista006 = async (file: File): Promise<void> => {
    try {
      const buffer = await file.arrayBuffer();
      const res = parsearExcelLista006(buffer, file.name, catalogoCompleto);
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

  const handleResolverEnBorrador = (idLinea: string, cantidades: number[]): boolean => {
    const lineaObj = lineasBorrador.find((l) => l.idLinea === idLinea);
    if (!lineaObj) return false;
    const resuelta = aplicarResolucionLinea(lineaObj, cantidades, catalogoCompleto);
    if (!resuelta) {
      onNotify(`La suma debe ser exactamente ${lineaObj.cantidadTotal} u.`, 'red');
      return false;
    }
    setLineasBorrador((prev) => prev.map((l) => (l.idLinea === idLinea ? resuelta : l)));
    onNotify(`${lineaObj.nombreProducto} confirmado.`, 'green');
    return true;
  };

  const resetFueraPlanilla = (): void => {
    setFpNombre('');
    setFpRubroSel('PILUSOS');
    setFpRubroCustom('');
    setFpPrecioNeto('');
    setFpTipoVar('unico');
    setFpCodigoUnico('');
    setFpCantidadTotal('1');
    setFpVariantes([
      { codigo: '', etiqueta: 'T1 (1-3 años)', cantidad: '' },
      { codigo: '', etiqueta: 'T2 (3-6 años)', cantidad: '' }
    ]);
    setFpGuardarEnCatalogo(true);
  };

  const handleAgregarFueraPlanilla = async (e: React.FormEvent<HTMLFormElement>): Promise<void> => {
    e.preventDefault();
    if (fpGuardando) return;
    const nombre = fpNombre.trim();
    const rubro = (fpRubroSel === '__OTRO__' ? fpRubroCustom : fpRubroSel).trim().toUpperCase();
    const precioNeto = round2(parseFloat(fpPrecioNeto.replace(',', '.')) || 0);

    if (!nombre) {
      onNotify('Ingresá el nombre del artículo.', 'amber');
      return;
    }
    if (!rubro) {
      onNotify('Ingresá el rubro del artículo.', 'amber');
      return;
    }

    const varsLimpias =
      fpTipoVar === 'unico'
        ? [{ codigo: fpCodigoUnico.trim(), etiqueta: 'Único', cantidad: fpCantidadTotal }]
        : fpVariantes.filter((v) => v.etiqueta.trim().length > 0);

    if (varsLimpias.length === 0) {
      onNotify('Cargá al menos una variante con nombre.', 'amber');
      return;
    }

    const cantidadesPorVar = varsLimpias.map((v) => Math.max(0, Math.floor(Number(v.cantidad) || 0)));
    const sumaVars = cantidadesPorVar.reduce((acc, q) => acc + q, 0);
    const totalInput = Math.max(0, Math.floor(Number(fpCantidadTotal) || 0));
    const cantidadTotal = fpTipoVar === 'unico' ? totalInput : sumaVars > 0 ? sumaVars : totalInput;

    if (cantidadTotal <= 0) {
      onNotify('Ingresá la cantidad pedida (mayor a 0).', 'amber');
      return;
    }

    let prodFinal: ProductoCatalogo;
    if (fpGuardarEnCatalogo) {
      setFpGuardando(true);
      const res = await crearProductoCatalogoAction({
        nombre,
        rubro,
        precioNeto,
        variantes: varsLimpias.map((v) => ({
          codigo: v.codigo.trim() || undefined,
          etiqueta: v.etiqueta.trim()
        }))
      });
      setFpGuardando(false);
      if (!res.ok || !res.data) {
        onNotify(res.error ?? 'No se pudo guardar en el catálogo.', 'red');
        return;
      }
      const created = res.data;
      onProductoCustomCreado(created);
      prodFinal = {
        idGrupo: created.idGrupo,
        nombre: created.nombre,
        rubro: created.rubro,
        precioNeto: created.precioNeto,
        variantes: created.variantes,
        esCustom: true,
        activo: created.activo,
        customId: created.id
      };
    } else {
      const adhocId = `ADHOC-${Date.now().toString(36).toUpperCase()}`;
      prodFinal = {
        idGrupo: adhocId,
        nombre,
        rubro,
        precioNeto,
        variantes: varsLimpias.map((v, idx) => ({
          codigo: v.codigo.trim().toUpperCase() || `${adhocId}-${idx + 1}`,
          etiqueta: v.etiqueta.trim(),
          aliases: [normalizarTexto(v.etiqueta)]
        }))
      };
    }

    let nuevaLinea: LineaPedido;
    if (prodFinal.variantes.length === 1) {
      nuevaLinea = resolverLinea(prodFinal, cantidadTotal, '');
    } else if (sumaVars === cantidadTotal && sumaVars > 0) {
      const desglose: DesgloseVariante[] = prodFinal.variantes.map((v, idx) => ({
        codigo: v.codigo,
        etiqueta: v.etiqueta,
        cantidad: cantidadesPorVar[idx] ?? 0
      }));
      nuevaLinea = {
        idLinea: prodFinal.idGrupo,
        idGrupo: prodFinal.idGrupo,
        nombreProducto: prodFinal.nombre,
        rubro: prodFinal.rubro,
        precioUnitarioNeto: prodFinal.precioNeto,
        cantidadTotal,
        notaCliente: desglose
          .filter((d) => d.cantidad > 0)
          .map((d) => `${d.cantidad} ${d.etiqueta}`)
          .join(' · '),
        estado: 'OK',
        requiereRevision: false,
        desglose
      };
    } else {
      nuevaLinea = resolverLinea(prodFinal, cantidadTotal, '');
    }

    setLineasBorrador((prev) => [...prev, nuevaLinea]);
    setCantidadesManuales((prev) => ({ ...prev, [prodFinal.idGrupo]: cantidadTotal }));
    resetFueraPlanilla();
    setMostrarFueraPlanilla(false);
    onNotify(
      fpGuardarEnCatalogo
        ? `${prodFinal.nombre} guardado en el catálogo y agregado al pedido.`
        : `${prodFinal.nombre} agregado solo a este pedido.`,
      'green'
    );
  };

  const renderVariantResolverBorrador = (l: LineaPedido, mode: 'pendiente' | 'edicion'): React.JSX.Element | null => {
    const prod = catalogoCompleto.find((p) => p.idGrupo === l.idGrupo);
    const variantes = prod ? prod.variantes : l.desglose;
    if (variantes.length < 2) return null;

    const key = mode === 'pendiente' ? `borr_${l.idLinea}` : `borr_ed_${l.idLinea}`;
    const defaultArr =
      mode === 'edicion'
        ? variantes.map((v) => String(l.desglose.find((d) => d.codigo === v.codigo)?.cantidad ?? 0))
        : variantes.map(() => '0');
    const cur = resolucionesManual[key] ?? defaultArr;
    const nums = variantes.map((_, idx) => Math.max(0, parseInt(cur[idx] ?? '0', 10) || 0));
    const suma = nums.reduce((acc, q) => acc + q, 0);
    const diff = l.cantidadTotal - suma;
    const esDos = variantes.length === 2;

    const setArr = (nextArr: string[]): void => {
      setResolucionesManual((prev) => ({ ...prev, [key]: nextArr }));
    };

    const confirmar = (): void => {
      const ok = handleResolverEnBorrador(l.idLinea, nums);
      if (ok && mode === 'edicion') setEditandoLineaId(null);
    };

    const handleInput = (vIdx: number, raw: string): void => {
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
        className={`rounded-xl border p-3.5 space-y-2.5 ${
          mode === 'pendiente'
            ? 'bg-white dark:bg-[#1B1B1F] border-red-200/70 dark:border-red-500/25'
            : 'bg-[#FAFAFA] dark:bg-[#1B1B1F] border-zinc-200 dark:border-[#2E2E35]'
        }`}
      >
        <div className="flex items-center justify-between text-xs">
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
            <div key={v.codigo} className="flex-1 min-w-[110px]">
              <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                {v.etiqueta} ({v.codigo})
              </label>
              <input
                type="number"
                min={0}
                max={l.cantidadTotal}
                value={cur[vIdx] ?? '0'}
                onWheel={(e) => e.currentTarget.blur()}
                onChange={(e) => handleInput(vIdx, e.target.value)}
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
              disabled={suma !== l.cantidadTotal}
              onClick={confirmar}
              className="px-3.5 py-1.5 rounded-full text-xs font-medium bg-zinc-900 hover:bg-zinc-800 text-white dark:bg-zinc-100 dark:hover:bg-white dark:text-zinc-950 disabled:opacity-40"
            >
              Confirmar
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

  const handleSubmit = async (): Promise<void> => {
    if (guardandoPedido) return;
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

  const fpPrecioNum = round2(parseFloat(fpPrecioNeto.replace(',', '.')) || 0);
  const fpPrecioConIva = round2(fpPrecioNum * 1.21);

  const hayDatosEnBorrador =
    lineasBorrador.length > 0 || form.cliente.trim().length > 0 || form.observaciones.trim().length > 0;

  const faltantesParaGuardar: string[] = [];
  if (!form.cliente.trim()) faltantesParaGuardar.push('Falta nombre del cliente');
  if (lineasBorrador.length === 0) faltantesParaGuardar.push('Cargá al menos 1 artículo');
  if (advertenciasExcel.length > 0) faltantesParaGuardar.push('Revisá las alertas del Excel');

  const msgReclamoBorrador = generarMensajeWhatsApp(
    form.cliente.trim(),
    proximoNumeroPedido,
    lineasBorrador,
    catalogoCompleto
  );
  const waUrlBorrador = buildWhatsAppUrl(form.whatsapp, msgReclamoBorrador);

  return (
    <div className="max-w-2xl mx-auto space-y-6">
      <datalist id={clientesListId}>
        {clientesConocidos.map((c) => (
          <option key={c.cliente} value={c.cliente} />
        ))}
      </datalist>

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

      <div className="flex items-center justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={onVolver}
            className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-800 hover:bg-zinc-100 dark:hover:text-[#F4F4F5] dark:hover:bg-[#1B1B1F] transition-colors"
            aria-label="Volver a pedidos"
          >
            <ArrowLeft className="w-4 h-4" strokeWidth={1.5} />
          </button>
          <div>
            <h1 className="text-lg font-semibold tracking-tight text-[#111111] dark:text-[#F4F4F5]">
              Nuevo pedido <span className="text-zinc-400 dark:text-zinc-500 font-normal tabular-nums">#{proximoNumeroPedido}</span>
            </h1>
            <p className="text-xs text-zinc-500 dark:text-zinc-400">Cargá el Excel Lista N° 006 o ingresá cantidades a mano.</p>
          </div>
        </div>

        {hayDatosEnBorrador && (
          confirmDescartar ? (
            <div className="inline-flex items-center gap-1.5 text-xs bg-red-50 border border-red-200 dark:bg-red-500/10 dark:border-red-500/25 rounded-full px-3 py-1">
              <span className="text-red-800 dark:text-red-200 font-medium">¿Descartar borrador?</span>
              <button
                type="button"
                onClick={handleDescartarBorrador}
                className="px-2 py-0.5 rounded-full bg-red-600 text-white font-semibold"
              >
                Sí
              </button>
              <button
                type="button"
                onClick={() => setConfirmDescartar(false)}
                className="px-1.5 py-0.5 text-zinc-600 dark:text-zinc-300"
              >
                No
              </button>
            </div>
          ) : (
            <button
              type="button"
              onClick={() => setConfirmDescartar(true)}
              className="text-xs text-zinc-400 hover:text-red-600 dark:text-zinc-500 dark:hover:text-red-400 transition-colors"
            >
              Descartar borrador
            </button>
          )
        )}
      </div>

      <section className="bg-white dark:bg-[#141417] rounded-2xl border border-[#EAEAEA] dark:border-[#26262B] shadow-[0_1px_2px_rgba(0,0,0,0.03)] p-6 space-y-4">
        <h2 className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 uppercase tracking-wider">1. Datos del pedido</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div className="sm:col-span-2 space-y-1.5">
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300">Cliente o negocio *</label>
              {clienteCoincidente && !form.whatsapp.trim() && (clienteCoincidente.whatsapp || clienteCoincidente.vendedor) && (
                <button
                  type="button"
                  onClick={() => {
                    setForm((prev) => ({
                      ...prev,
                      vendedor: clienteCoincidente.vendedor || prev.vendedor,
                      whatsapp: clienteCoincidente.whatsapp || prev.whatsapp
                    }));
                    onNotify(`Completados vendedor y WhatsApp de ${clienteCoincidente.cliente}.`, 'green');
                  }}
                  className="text-[11px] font-medium px-2.5 py-0.5 rounded-full bg-[#FCE7F3] text-[#111111] border border-[#F4B8D1] dark:bg-[#F4B6CE]/15 dark:text-[#FCE7F3] dark:border-[#F4B6CE]/30"
                >
                  Usar datos anteriores (vendedor y WhatsApp)
                </button>
              )}
            </div>
            <input
              type="text"
              list={clientesListId}
              value={form.cliente}
              onChange={(e) => setForm((prev) => ({ ...prev, cliente: e.target.value }))}
              placeholder="Ej. Casa Lorenzo, Pañalera Nuevo Mundo..."
              className="w-full px-3.5 py-2.5 text-sm bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1.5">Vendedor</label>
            <select
              value={form.vendedor}
              onChange={(e) => setForm((prev) => ({ ...prev, vendedor: e.target.value }))}
              className="w-full px-3.5 py-2.5 text-sm bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
            >
              {VENDEDORES_MIBI.map((v) => (
                <option key={v} value={v}>
                  {v}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1.5">WhatsApp del cliente</label>
            <input
              type="text"
              value={form.whatsapp}
              onChange={(e) => setForm((prev) => ({ ...prev, whatsapp: e.target.value }))}
              placeholder="Ej. 11-5606-9384"
              className="w-full px-3.5 py-2.5 text-sm bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70 tabular-nums"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1.5">Fecha pedido</label>
            <input
              type="date"
              value={form.fechaPedido}
              onChange={(e) => setForm((prev) => ({ ...prev, fechaPedido: e.target.value }))}
              className="w-full px-3.5 py-2.5 text-sm bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70 tabular-nums"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1.5">Despacho previsto</label>
            <input
              type="date"
              value={form.despachoPrevisto}
              onChange={(e) => setForm((prev) => ({ ...prev, despachoPrevisto: e.target.value }))}
              className="w-full px-3.5 py-2.5 text-sm bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70 tabular-nums"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1.5">Descuento (%)</label>
            <input
              type="number"
              min={0}
              max={100}
              step="0.5"
              disabled={!esAdmin}
              value={esAdmin ? form.descuentoPorc : 0}
              onWheel={(e) => e.currentTarget.blur()}
              onChange={(e) =>
                setForm((prev) => ({
                  ...prev,
                  descuentoPorc: Math.max(0, Math.min(100, Number(e.target.value) || 0))
                }))
              }
              className="w-full px-3.5 py-2.5 text-sm bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70 tabular-nums disabled:bg-zinc-50 dark:disabled:bg-[#141417] disabled:text-zinc-400 dark:disabled:text-zinc-500"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1.5">Observaciones</label>
            <input
              type="text"
              value={form.observaciones}
              onChange={(e) => setForm((prev) => ({ ...prev, observaciones: e.target.value }))}
              placeholder="Notas internas o entrega..."
              className="w-full px-3.5 py-2.5 text-sm bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
            />
          </div>
        </div>

        <div className="pt-3 border-t border-zinc-200/70 dark:border-[#26262B] space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-zinc-600 dark:text-zinc-300">Estado de pago inicial</span>
            {!esAdmin && (
              <span className="text-[11px] text-zinc-400 dark:text-zinc-500 inline-flex items-center gap-1">
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
                      ? 'bg-zinc-900 text-white border-zinc-900 dark:bg-zinc-100 dark:text-zinc-950 dark:border-zinc-100'
                      : 'bg-white text-zinc-700 border-zinc-200 hover:bg-zinc-50 dark:bg-[#1B1B1F] dark:text-zinc-300 dark:border-[#2E2E35] dark:hover:bg-[#24242B]'
                  }`}
                >
                  {tipo === 'Transferencia / Efectivo' ? 'Pagado / Seña' : tipo}
                </button>
              );
            })}
          </div>

          {esAdmin && form.pagoTipo === 'Transferencia / Efectivo' && (
            <div>
              <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1">
                Monto acreditado (vacío = total facturado)
              </label>
              <input
                type="text"
                value={form.montoAbonadoInput}
                onChange={(e) => setForm((prev) => ({ ...prev, montoAbonadoInput: e.target.value }))}
                placeholder={totalesBorrador.totalFacturado > 0 ? String(totalesBorrador.totalFacturado) : 'Total'}
                className="w-full px-3.5 py-2 text-sm bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500 tabular-nums"
              />
            </div>
          )}

          {esAdmin && form.pagoTipo === 'Cheque' && (
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1">Fecha recepción</label>
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
                  className="w-full px-3 py-2 text-xs bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1">Acreditación (+30d)</label>
                <input
                  type="date"
                  value={form.fechaAcreditacionCheque}
                  onChange={(e) => setForm((prev) => ({ ...prev, fechaAcreditacionCheque: e.target.value }))}
                  className="w-full px-3 py-2 text-xs bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-zinc-600 dark:text-zinc-300 mb-1">Detalle cheque</label>
                <input
                  type="text"
                  value={form.detalleCheque}
                  onChange={(e) => setForm((prev) => ({ ...prev, detalleCheque: e.target.value }))}
                  placeholder="Banco / N.º"
                  className="w-full px-3 py-2 text-xs bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                />
              </div>
            </div>
          )}
        </div>
      </section>

      <section className="bg-white dark:bg-[#141417] rounded-2xl border border-[#EAEAEA] dark:border-[#26262B] shadow-[0_1px_2px_rgba(0,0,0,0.03)] p-6 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <h2 className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 uppercase tracking-wider">
            2. Artículos (LISTA N° 006 y adicionales)
          </h2>
          <div className="inline-flex items-center p-0.5 rounded-full bg-zinc-100 dark:bg-[#1B1B1F] border border-zinc-200/70 dark:border-[#26262B]">
            <button
              type="button"
              onClick={() => setModoCargaManual(false)}
              className={`px-3 py-1 rounded-full text-xs font-medium transition-colors ${
                !modoCargaManual
                  ? 'bg-white text-[#111111] dark:bg-[#2A2A32] dark:text-white shadow-sm'
                  : 'text-zinc-500 dark:text-zinc-400'
              }`}
            >
              Subir Excel
            </button>
            <button
              type="button"
              onClick={() => setModoCargaManual(true)}
              className={`px-3 py-1 rounded-full text-xs font-medium transition-colors ${
                modoCargaManual
                  ? 'bg-white text-[#111111] dark:bg-[#2A2A32] dark:text-white shadow-sm'
                  : 'text-zinc-500 dark:text-zinc-400'
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
                ? 'border-[#F4B6CE] bg-pink-50/40 dark:border-[#F4B6CE] dark:bg-[#F4B6CE]/10'
                : 'border-zinc-200 hover:border-zinc-300 bg-[#FAFAFA] dark:bg-[#101014] dark:border-[#2E2E36] dark:hover:border-[#F4B6CE]/50'
            }`}
          >
            <FileSpreadsheet className="w-7 h-7 mx-auto text-zinc-400 dark:text-zinc-500" strokeWidth={1.5} />
            <p className="mt-2.5 text-sm font-medium text-[#111111] dark:text-[#F4F4F5]">
              {nombreArchivoCargado
                ? `Cargado: ${nombreArchivoCargado}`
                : 'Arrastrá la planilla LISTA N° 006 o hacé clic para elegir'}
            </p>
            <p className="mt-1 text-xs text-zinc-400 dark:text-zinc-500">Soporta .xlsx, .xls y .csv · Lectura exacta inmediata</p>
          </div>
        ) : (
          <div className="space-y-3">
            <div className="flex flex-wrap items-center gap-2">
              <input
                type="text"
                value={busquedaCatalogo}
                onChange={(e) => setBusquedaCatalogo(e.target.value)}
                placeholder="Filtrar catálogo por nombre, código o rubro..."
                className="flex-1 min-w-[200px] px-3.5 py-2 text-sm bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
              />
              {cargadosManualCount > 0 && (
                <button
                  type="button"
                  onClick={() => setSoloCargados((prev) => !prev)}
                  className={`px-3 py-2 rounded-xl text-xs font-medium border transition-colors shrink-0 tabular-nums ${
                    soloCargados
                      ? 'bg-zinc-900 text-white border-zinc-900 dark:bg-zinc-100 dark:text-zinc-950 dark:border-zinc-100'
                      : 'bg-white text-zinc-700 border-zinc-200 hover:bg-zinc-50 dark:bg-[#1B1B1F] dark:text-zinc-200 dark:border-[#2E2E35] dark:hover:bg-[#24242B]'
                  }`}
                >
                  Solo cargados ({cargadosManualCount})
                </button>
              )}
            </div>
            <div className="max-h-72 overflow-y-auto divide-y divide-zinc-200/60 dark:divide-[#26262B]/80 border border-zinc-200/80 dark:border-[#26262B] rounded-xl">
              {catalogoFiltradoManual.map((prod) => {
                const val = cantidadesManuales[prod.idGrupo] ?? 0;
                return (
                  <div key={prod.idGrupo} className="px-3.5 py-2.5 flex items-center justify-between gap-3">
                    <div className="min-w-0">
                      <div className="text-xs font-medium text-[#111111] dark:text-[#F4F4F5] truncate flex items-center gap-1.5">
                        <span className="truncate">{prod.nombre}</span>
                        {prod.esCustom && (
                          <span className="px-1.5 py-0.5 rounded bg-[#FCE7F3] text-[#111111] dark:bg-[#F4B6CE]/15 dark:text-[#FCE7F3] text-[10px] font-medium shrink-0">
                            Extra
                          </span>
                        )}
                      </div>
                      <div className="text-[11px] text-zinc-400 dark:text-zinc-500 tabular-nums">
                        {prod.variantes.map((v) => v.codigo).join(' / ')} · {prod.rubro} ·{' '}
                        {formatARSCompacto(prod.precioNeto)} neto
                      </div>
                    </div>
                    <input
                      type="number"
                      min={0}
                      value={val === 0 ? '' : val}
                      onWheel={(e) => e.currentTarget.blur()}
                      onChange={(e) => handleCambioCantidadManual(prod, e.target.value)}
                      placeholder="0"
                      className="w-20 px-2.5 py-1.5 text-sm text-right bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg tabular-nums placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
                    />
                  </div>
                );
              })}
            </div>
          </div>
        )}

        <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
          <div className="flex flex-wrap items-center gap-3">
            <button
              type="button"
              onClick={() => setMostrarFueraPlanilla((prev) => !prev)}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-full text-xs font-medium bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 text-[#111111] dark:bg-[#1B1B1F] dark:hover:bg-[#24242B] dark:border-[#2E2E35] dark:text-[#F4F4F5] transition-colors"
              title="Ej. Pilusos o artículos nuevos"
            >
              <Plus className="w-3.5 h-3.5" strokeWidth={1.75} />
              Producto fuera de planilla
            </button>
            {lineasBorrador.length === 0 && (
              <button
                type="button"
                onClick={handleCargarEjemploLista006}
                className="text-xs text-zinc-400 hover:text-[#111111] dark:text-zinc-500 dark:hover:text-[#F4F4F5] underline underline-offset-2 transition-colors"
              >
                Probar con planilla de ejemplo
              </button>
            )}
          </div>
          {lineasBorrador.length > 0 && (
            confirmVaciarArticulos ? (
              <div className="inline-flex items-center gap-1.5 text-xs bg-red-50 border border-red-200 dark:bg-red-500/10 dark:border-red-500/25 rounded-full px-3 py-1">
                <span className="text-red-800 dark:text-red-200">¿Vaciar {lineasBorrador.length} ítems?</span>
                <button
                  type="button"
                  onClick={() => {
                    setLineasBorrador([]);
                    setCantidadesManuales({});
                    setAdvertenciasExcel([]);
                    setNombreArchivoCargado('');
                    setSoloCargados(false);
                    setConfirmVaciarArticulos(false);
                  }}
                  className="px-2 py-0.5 rounded-full bg-red-600 text-white font-semibold"
                >
                  Sí
                </button>
                <button
                  type="button"
                  onClick={() => setConfirmVaciarArticulos(false)}
                  className="px-1.5 py-0.5 text-zinc-600 dark:text-zinc-300"
                >
                  No
                </button>
              </div>
            ) : (
              <button
                type="button"
                onClick={() => setConfirmVaciarArticulos(true)}
                className="text-xs text-zinc-400 hover:text-red-600 dark:hover:text-red-400 transition-colors"
              >
                Vaciar artículos
              </button>
            )
          )}
        </div>

        {mostrarFueraPlanilla && (
          <form
            onSubmit={(e) => void handleAgregarFueraPlanilla(e)}
            className="bg-[#FAFAFA] dark:bg-[#1B1B1F] rounded-2xl border border-zinc-200/80 dark:border-[#26262B] p-4 space-y-3.5 text-xs"
          >
            <div className="flex items-center justify-between">
              <div>
                <div className="font-semibold text-[#111111] dark:text-[#F4F4F5]">Agregar producto fuera de planilla</div>
                <div className="text-[11px] text-zinc-500 dark:text-zinc-400">Ej. Pilusos o artículos nuevos</div>
              </div>
              <button
                type="button"
                onClick={() => setMostrarFueraPlanilla(false)}
                className="p-1 text-zinc-400 hover:text-zinc-800 dark:hover:text-[#F4F4F5]"
              >
                <X className="w-4 h-4" strokeWidth={1.5} />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Nombre del producto *</label>
                <input
                  type="text"
                  value={fpNombre}
                  onChange={(e) => setFpNombre(e.target.value)}
                  placeholder="Ej. Piluso Shark"
                  className="w-full px-3 py-2 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                />
              </div>

              <div>
                <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Rubro *</label>
                <div className="flex gap-2">
                  <select
                    value={fpRubroSel}
                    onChange={(e) => setFpRubroSel(e.target.value)}
                    className="flex-1 px-3 py-2 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl"
                  >
                    {rubrosDisponibles.map((r) => (
                      <option key={r} value={r}>
                        {r}
                      </option>
                    ))}
                    <option value="__OTRO__">Otro rubro...</option>
                  </select>
                  {fpRubroSel === '__OTRO__' && (
                    <input
                      type="text"
                      value={fpRubroCustom}
                      onChange={(e) => setFpRubroCustom(e.target.value)}
                      placeholder="Ej. PILUSOS"
                      className="w-28 px-3 py-2 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl uppercase placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                    />
                  )}
                </div>
              </div>

              <div>
                <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Precio neto unitario ($)</label>
                <input
                  type="number"
                  min={0}
                  step="0.01"
                  value={fpPrecioNeto}
                  onWheel={(e) => e.currentTarget.blur()}
                  onChange={(e) => setFpPrecioNeto(e.target.value)}
                  placeholder="0.00"
                  className="w-full px-3 py-2 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                />
                <div className="text-[11px] text-zinc-400 dark:text-zinc-500 mt-1 tabular-nums">
                  Con IVA 21%: {formatARS(fpPrecioConIva)}
                </div>
              </div>

              <div>
                <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Variantes</label>
                <div className="grid grid-cols-2 gap-1.5">
                  <button
                    type="button"
                    onClick={() => setFpTipoVar('unico')}
                    className={`py-2 px-3 rounded-xl border font-medium ${
                      fpTipoVar === 'unico'
                        ? 'bg-zinc-900 text-white border-zinc-900 dark:bg-zinc-100 dark:text-zinc-950 dark:border-zinc-100'
                        : 'bg-white text-zinc-700 border-zinc-200 dark:bg-[#141417] dark:text-zinc-300 dark:border-[#2E2E35]'
                    }`}
                  >
                    Único
                  </button>
                  <button
                    type="button"
                    onClick={() => setFpTipoVar('multiple')}
                    className={`py-2 px-3 rounded-xl border font-medium ${
                      fpTipoVar === 'multiple'
                        ? 'bg-zinc-900 text-white border-zinc-900 dark:bg-zinc-100 dark:text-zinc-950 dark:border-zinc-100'
                        : 'bg-white text-zinc-700 border-zinc-200 dark:bg-[#141417] dark:text-zinc-300 dark:border-[#2E2E35]'
                    }`}
                  >
                    Con talles o colores
                  </button>
                </div>
              </div>
            </div>

            {fpTipoVar === 'unico' ? (
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Código (opcional)</label>
                  <input
                    type="text"
                    value={fpCodigoUnico}
                    onChange={(e) => setFpCodigoUnico(e.target.value)}
                    placeholder="Vacío → automático"
                    className="w-full px-3 py-2 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums uppercase placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">Cantidad pedida *</label>
                  <input
                    type="number"
                    min={1}
                    value={fpCantidadTotal}
                    onWheel={(e) => e.currentTarget.blur()}
                    onChange={(e) => setFpCantidadTotal(e.target.value)}
                    className="w-full px-3 py-2 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl tabular-nums"
                  />
                </div>
              </div>
            ) : (
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-zinc-500 dark:text-zinc-400">
                    Opciones (indicá cantidad en cada una o dejá vacío para definir después)
                  </span>
                  <button
                    type="button"
                    onClick={() =>
                      setFpVariantes((prev) => [...prev, { codigo: '', etiqueta: '', cantidad: '' }])
                    }
                    className="text-[11px] font-medium text-zinc-800 dark:text-zinc-200 hover:underline inline-flex items-center gap-1"
                  >
                    <Plus className="w-3 h-3" strokeWidth={1.75} /> Agregar variante
                  </button>
                </div>
                {fpVariantes.map((v, idx) => (
                  <div key={idx} className="flex items-center gap-2">
                    <input
                      type="text"
                      value={v.codigo}
                      onChange={(e) =>
                        setFpVariantes((prev) =>
                          prev.map((item, i) => (i === idx ? { ...item, codigo: e.target.value } : item))
                        )
                      }
                      placeholder="Cód. (opc.)"
                      className="w-24 px-2.5 py-1.5 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg tabular-nums uppercase placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                    />
                    <input
                      type="text"
                      value={v.etiqueta}
                      onChange={(e) =>
                        setFpVariantes((prev) =>
                          prev.map((item, i) => (i === idx ? { ...item, etiqueta: e.target.value } : item))
                        )
                      }
                      placeholder={`Talle o color #${idx + 1}`}
                      className="flex-1 px-2.5 py-1.5 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                    />
                    <input
                      type="number"
                      min={0}
                      value={v.cantidad}
                      onWheel={(e) => e.currentTarget.blur()}
                      onChange={(e) =>
                        setFpVariantes((prev) =>
                          prev.map((item, i) => (i === idx ? { ...item, cantidad: e.target.value } : item))
                        )
                      }
                      placeholder="Cant."
                      className="w-20 px-2.5 py-1.5 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg text-right tabular-nums placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                    />
                    {fpVariantes.length > 2 && (
                      <button
                        type="button"
                        onClick={() => setFpVariantes((prev) => prev.filter((_, i) => i !== idx))}
                        className="p-1.5 text-zinc-400 hover:text-red-600 dark:hover:text-red-400"
                        aria-label="Quitar variante"
                      >
                        <Trash2 className="w-3.5 h-3.5" strokeWidth={1.5} />
                      </button>
                    )}
                  </div>
                ))}

                {fpVariantes.every((v) => !(parseInt(v.cantidad, 10) > 0)) && (
                  <div>
                    <label className="block text-[11px] text-zinc-500 dark:text-zinc-400 mb-1">
                      Cantidad total pendiente de definir
                    </label>
                    <input
                      type="number"
                      min={1}
                      value={fpCantidadTotal}
                      onWheel={(e) => e.currentTarget.blur()}
                      onChange={(e) => setFpCantidadTotal(e.target.value)}
                      className="w-32 px-3 py-1.5 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg tabular-nums"
                    />
                  </div>
                )}
              </div>
            )}

            <label className="flex items-center gap-2 pt-1 cursor-pointer select-none text-zinc-700 dark:text-zinc-300">
              <input
                type="checkbox"
                checked={fpGuardarEnCatalogo}
                onChange={(e) => setFpGuardarEnCatalogo(e.target.checked)}
                className="w-4 h-4 rounded border-zinc-300 dark:border-zinc-600 text-zinc-900 accent-zinc-900 dark:accent-[#F4B6CE]"
              />
              <span>Guardar en el catálogo para próximos pedidos</span>
            </label>

            <div className="flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setMostrarFueraPlanilla(false)}
                className="px-3.5 py-2 rounded-full border border-zinc-200 dark:border-[#2E2E35] bg-white dark:bg-[#141417] text-zinc-600 dark:text-zinc-300"
              >
                Cancelar
              </button>
              <button
                type="submit"
                disabled={fpGuardando}
                className="px-4 py-2 rounded-full bg-zinc-900 text-white font-medium hover:bg-zinc-800 dark:bg-zinc-100 dark:text-zinc-950 dark:hover:bg-white disabled:opacity-50"
              >
                {fpGuardando ? 'Guardando...' : 'Agregar al pedido'}
              </button>
            </div>
          </form>
        )}
      </section>

      {advertenciasExcel.length > 0 && (
        <section className="bg-red-50/70 border border-red-200 dark:bg-red-500/10 dark:border-red-500/25 rounded-2xl p-5 space-y-3">
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-xs font-semibold text-red-800 dark:text-red-300">
              <AlertTriangle className="w-4 h-4 text-red-600 dark:text-red-400" strokeWidth={1.75} />
              <span>Errores bloqueantes detectados en el Excel ({advertenciasExcel.length})</span>
            </div>
            <button
              type="button"
              onClick={() => setAdvertenciasExcel([])}
              className="text-[11px] font-medium text-red-700 dark:text-red-300 hover:underline"
            >
              Confirmar y descartar avisos
            </button>
          </div>
          <ul className="space-y-1 text-xs text-red-700 dark:text-red-300/90">
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
        <section className="bg-red-50/60 border border-red-200/80 dark:bg-red-500/10 dark:border-red-500/25 rounded-2xl p-5 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="text-xs font-semibold text-red-800 dark:text-red-300">
              Falta definir talle o color en {totalesBorrador.faltantes}{' '}
              {totalesBorrador.faltantes === 1 ? 'artículo' : 'artículos'}
            </div>

            <div className="inline-flex items-center gap-1.5">
              {waUrlBorrador ? (
                <>
                  <a
                    href={waUrlBorrador}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors"
                  >
                    <MessageCircle className="w-3.5 h-3.5" strokeWidth={1.75} />
                    Enviar por WhatsApp
                  </a>
                  <button
                    type="button"
                    title="Copiar mensaje"
                    aria-label="Copiar mensaje para WhatsApp"
                    onClick={() =>
                      onCopyText('borrador-wsp', msgReclamoBorrador, 'Mensaje de consulta copiado.')
                    }
                    className="p-2 rounded-full bg-white border border-red-200 text-red-700 hover:bg-red-50 dark:bg-[#1B1B1F] dark:border-red-500/30 dark:text-red-300"
                  >
                    {copiadoId === 'borrador-wsp' ? (
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
                      'borrador-wsp',
                      msgReclamoBorrador,
                      'Mensaje de consulta copiado para WhatsApp.'
                    )
                  }
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-white border border-red-200 text-red-700 hover:bg-red-50 dark:bg-[#1B1B1F] dark:border-red-500/30 dark:text-red-300 dark:hover:bg-red-500/15"
                >
                  {copiadoId === 'borrador-wsp' ? (
                    <Check className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400" strokeWidth={1.75} />
                  ) : (
                    <MessageCircle className="w-3.5 h-3.5" strokeWidth={1.5} />
                  )}
                  {copiadoId === 'borrador-wsp' ? 'WhatsApp copiado' : 'Copiar para WhatsApp'}
                </button>
              )}
            </div>
          </div>

          <div className="space-y-2.5">
            {lineasBorrador
              .filter((l) => l.requiereRevision || l.estado !== 'OK')
              .map((l) => renderVariantResolverBorrador(l, 'pendiente'))}
          </div>
        </section>
      )}

      {(lineasBorrador.length > 0 || hayDatosEnBorrador) && (
        <section className="bg-white dark:bg-[#141417] rounded-2xl border border-[#EAEAEA] dark:border-[#26262B] shadow-[0_1px_2px_rgba(0,0,0,0.03)] p-6 space-y-4">
          {lineasBorrador.length > 0 && (
            <>
              <div className="flex items-center justify-between">
                <h2 className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 uppercase tracking-wider">
                  3. Resumen y liquidación ({lineasBorrador.length}{' '}
                  {lineasBorrador.length === 1 ? 'ítem' : 'ítems'} · {totalesBorrador.totalUnidades} u.)
                </h2>
              </div>

              <div className="max-h-72 overflow-y-auto divide-y divide-zinc-200/60 dark:divide-[#26262B]/80 border border-zinc-200/70 dark:border-[#26262B] rounded-xl">
                {lineasBorrador.map((l) => {
                  const prod = catalogoCompleto.find((p) => p.idGrupo === l.idGrupo);
                  const variantes = prod ? prod.variantes : l.desglose;
                  const puedeEditarTalles = !l.requiereRevision && variantes.length >= 2;
                  return (
                    <div key={l.idLinea} className="px-3.5 py-2.5 space-y-2 text-xs">
                      <div className="flex items-center justify-between gap-3">
                        <div className="min-w-0">
                          <div className="font-medium text-[#111111] dark:text-[#F4F4F5] truncate">
                            {l.nombreProducto}
                          </div>
                          <div className="text-[11px] text-zinc-400 dark:text-zinc-500 tabular-nums">
                            {l.requiereRevision
                              ? `Pendiente definir (${l.cantidadTotal} u.)`
                              : l.desglose
                                  .filter((d) => d.cantidad > 0)
                                  .map((d) => `${d.codigo} ${d.etiqueta}: ${d.cantidad}`)
                                  .join(' · ')}
                          </div>
                        </div>
                        <div className="flex items-center gap-2.5 shrink-0 tabular-nums">
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
                          <span className="font-semibold text-[#111111] dark:text-[#F4F4F5]">
                            {formatARS(round2(l.precioUnitarioNeto * l.cantidadTotal))}
                          </span>
                          <button
                            type="button"
                            onClick={() => handleEliminarLineaBorrador(l.idGrupo)}
                            className="p-1 text-zinc-400 hover:text-red-600 dark:hover:text-red-400"
                            aria-label="Quitar artículo"
                          >
                            <X className="w-3.5 h-3.5" strokeWidth={1.5} />
                          </button>
                        </div>
                      </div>

                      {editandoLineaId === l.idLinea && renderVariantResolverBorrador(l, 'edicion')}
                    </div>
                  );
                })}
              </div>

              <div className="pt-2 space-y-1.5 text-xs tabular-nums">
                <div className="flex justify-between text-zinc-500 dark:text-zinc-400">
                  <span>Subtotal Neto</span>
                  <span>{formatARS(totalesBorrador.subtotalNeto)}</span>
                </div>
                {totalesBorrador.montoDescuento > 0 && (
                  <div className="flex justify-between text-zinc-500 dark:text-zinc-400">
                    <span>Descuento ({form.descuentoPorc}%)</span>
                    <span>-{formatARS(totalesBorrador.montoDescuento)}</span>
                  </div>
                )}
                <div className="flex justify-between text-zinc-500 dark:text-zinc-400">
                  <span>Base Imponible (Total B)</span>
                  <span>{formatARS(totalesBorrador.baseImponible)}</span>
                </div>
                <div className="flex justify-between text-zinc-500 dark:text-zinc-400">
                  <span>IVA 21%</span>
                  <span>{formatARS(totalesBorrador.iva21)}</span>
                </div>
                <div className="flex justify-between text-sm font-semibold text-[#111111] dark:text-[#F4F4F5] pt-2 border-t border-zinc-200/70 dark:border-[#26262B]">
                  <span>Total Facturado</span>
                  <span>{formatARS(totalesBorrador.totalFacturado)}</span>
                </div>
              </div>
            </>
          )}

          {faltantesParaGuardar.length > 0 && (
            <p className="text-xs text-amber-700 dark:text-amber-400 text-center">
              Para guardar: {faltantesParaGuardar.join(' · ')}.
            </p>
          )}

          <button
            type="button"
            disabled={guardandoPedido || faltantesParaGuardar.length > 0}
            onClick={() => void handleSubmit()}
            className="w-full py-3 rounded-full text-sm font-medium bg-zinc-900 hover:bg-zinc-800 text-white dark:bg-zinc-100 dark:hover:bg-white dark:text-zinc-950 transition-colors disabled:opacity-45 disabled:cursor-not-allowed"
          >
            {guardandoPedido ? 'Guardando...' : 'Guardar pedido'}
          </button>
        </section>
      )}
    </div>
  );
}
