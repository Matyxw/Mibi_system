'use client';

import React, { useEffect, useId, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  ArrowLeft,
  Check,
  ChevronDown,
  ChevronUp,
  Copy,
  Edit3,
  FileSpreadsheet,
  Info,
  Lock,
  MessageCircle,
  Minus,
  Plus,
  RotateCcw,
  Search,
  Trash2,
  X
} from 'lucide-react';
import { crearProductoCatalogoAction } from '@/app/actions/catalogo-actions';
import { VENDEDORES_MIBI } from '@/lib/catalogo-mibi';
import {
  aplicarLineaAlBorrador,
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
  AvisoPrecioExcel,
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
  const fueraPlanillaRef = useRef<HTMLFormElement | null>(null);

  // Estado formulario "+ Producto fuera de planilla"
  const [mostrarFueraPlanilla, setMostrarFueraPlanilla] = useState<boolean>(false);
  const [fpNombre, setFpNombre] = useState<string>('');
  const [fpRubroSel, setFpRubroSel] = useState<string>('PILUSOS');
  const [fpRubroCustom, setFpRubroCustom] = useState<string>('');
  const [fpPrecioNeto, setFpPrecioNeto] = useState<string>('');
  const [fpPrecioExcelHint, setFpPrecioExcelHint] = useState<number | null>(null);
  const [fpTipoVar, setFpTipoVar] = useState<'unico' | 'multiple'>('unico');
  const [fpCodigoUnico, setFpCodigoUnico] = useState<string>('');
  const [fpCantidadTotal, setFpCantidadTotal] = useState<string>('1');
  const [fpVariantes, setFpVariantes] = useState<CustomVarianteDraft[]>([
    { codigo: '', etiqueta: 'T1 (1-3 años)', cantidad: '' },
    { codigo: '', etiqueta: 'T2 (3-6 años)', cantidad: '' }
  ]);
  const [fpGuardarEnCatalogo, setFpGuardarEnCatalogo] = useState<boolean>(true);
  const [fpGuardando, setFpGuardando] = useState<boolean>(false);
  const [alertaCreandoId, setAlertaCreandoId] = useState<string | null>(null);
  const [alertaCreandoNota, setAlertaCreandoNota] = useState<string>('');

  // Centro de resolución de alertas del Excel
  const [avisosPrecio, setAvisosPrecio] = useState<AvisoPrecioExcel[]>([]);
  const [mostrarDetallePrecios, setMostrarDetallePrecios] = useState<boolean>(false);
  const [alertasIgnoradas, setAlertasIgnoradas] = useState<AdvertenciaExcel[]>([]);
  const [confirmIgnorarTodas, setConfirmIgnorarTodas] = useState<boolean>(false);
  const [corrigiendoAlertaId, setCorrigiendoAlertaId] = useState<string | null>(null);
  const [busquedaCorreccion, setBusquedaCorreccion] = useState<string>('');
  const [cantidadesCorreccion, setCantidadesCorreccion] = useState<Record<string, string>>({});

  useEffect(() => {
    if (!confirmIgnorarTodas) return;
    const t = window.setTimeout(() => setConfirmIgnorarTodas(false), 4000);
    return () => window.clearTimeout(t);
  }, [confirmIgnorarTodas]);

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

  // Ajuste directo en Paso 3: actualiza lineasBorrador y cantidadesManuales sin invocar sincronizarManualConBorrador
  const handleAjusteRapidoLineaPaso3 = (idGrupo: string, nuevaCantidad: number): void => {
    const qty = Math.max(1, Math.trunc(nuevaCantidad));
    setLineasBorrador((prev) =>
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
    setCantidadesManuales((prev) => ({ ...prev, [idGrupo]: qty }));
  };

  const handleAjusteVariantePaso3 = (
    idGrupo: string,
    codigoVar: string,
    siguienteVarQty: number
  ): void => {
    const varQty = Math.max(0, Math.trunc(siguienteVarQty));
    const lineaActual = lineasBorrador.find((l) => l.idGrupo === idGrupo);
    if (!lineaActual) return;
    const prod = catalogoCompleto.find((p) => p.idGrupo === idGrupo);
    const variantesBase = prod && prod.variantes.length > 0 ? prod.variantes : lineaActual.desglose;

    const nuevoDesglose = variantesBase.map((v) => {
      const prevCant = lineaActual.desglose.find((d) => d.codigo === v.codigo)?.cantidad ?? 0;
      return {
        codigo: v.codigo,
        etiqueta: v.etiqueta,
        cantidad: v.codigo === codigoVar ? varQty : prevCant
      };
    });

    const nuevaSuma = nuevoDesglose.reduce((acc, d) => acc + d.cantidad, 0);
    if (nuevaSuma <= 0) return;

    setLineasBorrador((prev) =>
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
    setCantidadesManuales((prev) => ({ ...prev, [idGrupo]: nuevaSuma }));
  };

  const limpiarEstadoAlertas = (): void => {
    setAdvertenciasExcel([]);
    setAvisosPrecio([]);
    setMostrarDetallePrecios(false);
    setAlertasIgnoradas([]);
    setConfirmIgnorarTodas(false);
    setCorrigiendoAlertaId(null);
    setBusquedaCorreccion('');
    setCantidadesCorreccion({});
    setAlertaCreandoId(null);
    setAlertaCreandoNota('');
    setFpPrecioExcelHint(null);
  };

  const handleDescartarBorrador = (): void => {
    setForm(crearFormInicial());
    setLineasBorrador([]);
    setCantidadesManuales({});
    limpiarEstadoAlertas();
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
      setAvisosPrecio(res.avisosPrecio ? [...res.avisosPrecio] : []);
      setMostrarDetallePrecios(false);
      setAlertasIgnoradas([]);
      setConfirmIgnorarTodas(false);
      setCorrigiendoAlertaId(null);
      setBusquedaCorreccion('');
      setCantidadesCorreccion({});
      setAlertaCreandoId(null);
      setAlertaCreandoNota('');
      setFpPrecioExcelHint(null);
      setNombreArchivoCargado(file.name);
      if (res.advertencias.length > 0) {
        onNotify(
          `Se detectaron ${res.advertencias.length} fila(s) para revisar en el centro de alertas antes de guardar.`,
          'amber'
        );
      } else if (res.avisosPrecio && res.avisosPrecio.length > 0) {
        onNotify(
          `${res.lineas.length} artículos leídos de ${file.name} (se aplicaron precios vigentes).`,
          'green'
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
    limpiarEstadoAlertas();
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
    setFpPrecioExcelHint(null);
    setFpTipoVar('unico');
    setFpCodigoUnico('');
    setFpCantidadTotal('1');
    setFpVariantes([
      { codigo: '', etiqueta: 'T1 (1-3 años)', cantidad: '' },
      { codigo: '', etiqueta: 'T2 (3-6 años)', cantidad: '' }
    ]);
    setFpGuardarEnCatalogo(true);
    setAlertaCreandoId(null);
    setAlertaCreandoNota('');
  };

  const handleIniciarCrearDesdeAlerta = (adv: AdvertenciaExcel): void => {
    setFpNombre(adv.descripcionExcel || '');
    setFpRubroSel('PILUSOS');
    setFpRubroCustom('');
    // Seguridad contable: el precio neto siempre nace vacío para confirmación humana explícita
    setFpPrecioNeto('');
    setFpPrecioExcelHint(adv.precioExcelReferencia);
    if (adv.codigosDetectados.length >= 2) {
      setFpTipoVar('multiple');
      setFpVariantes(
        adv.codigosDetectados.map((cod, idx) => ({
          codigo: cod,
          etiqueta: idx === 0 ? 'T1 / Opción 1' : idx === 1 ? 'T2 / Opción 2' : `Opción ${idx + 1}`,
          cantidad: ''
        }))
      );
      setFpCantidadTotal(String(adv.cantidadDetectada ?? 1));
    } else {
      setFpTipoVar('unico');
      setFpCodigoUnico(adv.codigosDetectados[0] ?? adv.codigo ?? '');
      setFpCantidadTotal(String(adv.cantidadDetectada ?? 1));
    }
    setFpGuardarEnCatalogo(true);
    setAlertaCreandoId(adv.id);
    setAlertaCreandoNota(adv.notaDetectada);
    setCorrigiendoAlertaId(null);
    setMostrarFueraPlanilla(true);
    window.requestAnimationFrame(() => {
      fueraPlanillaRef.current?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    });
  };

  const handleVincularAlertaConProducto = (adv: AdvertenciaExcel, prod: ProductoCatalogo): void => {
    const qty = adv.cantidadDetectada ?? 0;
    if (qty <= 0) {
      onNotify('Primero corregí la cantidad de la fila.', 'amber');
      return;
    }
    const nextLineas = aplicarLineaAlBorrador(lineasBorrador, prod, qty, adv.notaDetectada);
    const lineaFinal = nextLineas.find((l) => l.idGrupo === prod.idGrupo);
    setLineasBorrador(nextLineas);
    if (lineaFinal) {
      setCantidadesManuales((prev) => ({ ...prev, [prod.idGrupo]: lineaFinal.cantidadTotal }));
    }
    const nextAlertas = advertenciasExcel.filter((a) => a.id !== adv.id);
    setAdvertenciasExcel(nextAlertas);
    setCorrigiendoAlertaId(null);
    setBusquedaCorreccion('');
    if (nextAlertas.length === 0) {
      onNotify(`Sumado ${qty} u. a ${prod.nombre}. Todas las alertas resueltas.`, 'green');
    } else {
      onNotify(`Sumado ${qty} u. a ${prod.nombre}.`, 'green');
    }
  };

  const handleIgnorarAlerta = (adv: AdvertenciaExcel): void => {
    const nextAlertas = advertenciasExcel.filter((a) => a.id !== adv.id);
    setAdvertenciasExcel(nextAlertas);
    setAlertasIgnoradas((prev) => [...prev, adv]);
    if (corrigiendoAlertaId === adv.id) setCorrigiendoAlertaId(null);
    if (alertaCreandoId === adv.id) resetFueraPlanilla();
    if (nextAlertas.length === 0) {
      onNotify(`Fila ${adv.fila} ignorada. No quedan alertas pendientes.`, 'neutral');
    } else {
      onNotify(`Fila ${adv.fila} ignorada (podés restaurarla desde la lista).`, 'neutral');
    }
  };

  const handleIgnorarTodasAlertas = (): void => {
    if (advertenciasExcel.length === 0) return;
    setAlertasIgnoradas((prev) => [...prev, ...advertenciasExcel]);
    setAdvertenciasExcel([]);
    setConfirmIgnorarTodas(false);
    setCorrigiendoAlertaId(null);
    onNotify('Alertas ignoradas para este pedido. Podés restaurarlas si necesitás revisar.', 'neutral');
  };

  const handleRestaurarAlertasIgnoradas = (): void => {
    if (alertasIgnoradas.length === 0) return;
    const count = alertasIgnoradas.length;
    setAdvertenciasExcel((prev) => [...prev, ...alertasIgnoradas]);
    setAlertasIgnoradas([]);
    onNotify(`Se restauraron ${count} ${count === 1 ? 'fila ignorada' : 'filas ignoradas'}.`, 'amber');
  };

  const handleCorregirCantidadAlerta = (adv: AdvertenciaExcel): void => {
    const raw = (cantidadesCorreccion[adv.id] ?? '').trim();
    if (!/^\d+$/.test(raw)) {
      onNotify('Ingresá una cantidad entera válida mayor a 0.', 'amber');
      return;
    }
    const nuevaQty = parseInt(raw, 10);
    if (!Number.isInteger(nuevaQty) || nuevaQty <= 0) {
      onNotify('La cantidad debe ser al menos 1 unidad.', 'amber');
      return;
    }

    const prodExistente = adv.productoExistenteIdGrupo
      ? catalogoCompleto.find((p) => p.idGrupo === adv.productoExistenteIdGrupo)
      : undefined;

    if (prodExistente) {
      const nextLineas = aplicarLineaAlBorrador(lineasBorrador, prodExistente, nuevaQty, adv.notaDetectada);
      const lineaFinal = nextLineas.find((l) => l.idGrupo === prodExistente.idGrupo);
      setLineasBorrador(nextLineas);
      if (lineaFinal) {
        setCantidadesManuales((prev) => ({ ...prev, [prodExistente.idGrupo]: lineaFinal.cantidadTotal }));
      }
      const nextAlertas = advertenciasExcel.filter((a) => a.id !== adv.id);
      setAdvertenciasExcel(nextAlertas);
      setCantidadesCorreccion((prev) => {
        const copy = { ...prev };
        delete copy[adv.id];
        return copy;
      });
      if (nextAlertas.length === 0) {
        onNotify(`${prodExistente.nombre} (${nuevaQty} u.) sumado. Todas las alertas resueltas.`, 'green');
      } else {
        onNotify(`${prodExistente.nombre} (${nuevaQty} u.) sumado al pedido.`, 'green');
      }
      return;
    }

    // Si el código tampoco existía en catálogo, actualizamos la alerta a CODIGO_DESCONOCIDO con la nueva cantidad
    setAdvertenciasExcel((prev) =>
      prev.map((a) =>
        a.id === adv.id
          ? {
              ...a,
              tipo: 'CODIGO_DESCONOCIDO',
              cantidadDetectada: nuevaQty,
              mensaje: `Código desconocido (${a.codigo}) con cantidad ${nuevaQty}`
            }
          : a
      )
    );
    setCantidadesCorreccion((prev) => {
      const copy = { ...prev };
      delete copy[adv.id];
      return copy;
    });
    onNotify('Cantidad corregida. Ahora podés crear o vincular este producto.', 'amber');
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

    const notaInicial = alertaCreandoId ? alertaCreandoNota : '';
    let nuevaLinea: LineaPedido;
    if (prodFinal.variantes.length === 1) {
      nuevaLinea = resolverLinea(prodFinal, cantidadTotal, notaInicial);
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
      nuevaLinea = resolverLinea(prodFinal, cantidadTotal, notaInicial);
    }

    let nextLineas = [...lineasBorrador, nuevaLinea];
    let resueltasExtra = 0;
    let nextAlertas = advertenciasExcel;

    if (alertaCreandoId) {
      const codigosNuevos = new Set<string>(
        [...varsLimpias.map((v) => v.codigo.trim().toUpperCase()), ...prodFinal.variantes.map((v) => v.codigo.trim().toUpperCase())].filter(
          Boolean
        )
      );
      const pendientes: AdvertenciaExcel[] = [];
      for (const adv of advertenciasExcel) {
        if (adv.id === alertaCreandoId) continue;
        const coincideMismoCodigo =
          adv.tipo === 'CODIGO_DESCONOCIDO' &&
          (adv.cantidadDetectada ?? 0) > 0 &&
          adv.codigosDetectados.some((c) => codigosNuevos.has(c.toUpperCase()));
        if (coincideMismoCodigo) {
          nextLineas = aplicarLineaAlBorrador(
            nextLineas,
            prodFinal,
            adv.cantidadDetectada ?? 0,
            adv.notaDetectada
          );
          resueltasExtra++;
        } else {
          pendientes.push(adv);
        }
      }
      nextAlertas = pendientes;
      setAdvertenciasExcel(nextAlertas);
    }

    const lineaFinal = nextLineas.find((l) => l.idGrupo === prodFinal.idGrupo) ?? nuevaLinea;
    setLineasBorrador(nextLineas);
    setCantidadesManuales((prev) => ({ ...prev, [prodFinal.idGrupo]: lineaFinal.cantidadTotal }));
    const habiaAlerta = Boolean(alertaCreandoId);
    resetFueraPlanilla();
    setMostrarFueraPlanilla(false);

    if (habiaAlerta) {
      const sufijoAuto = resueltasExtra > 0 ? ` (+${resueltasExtra} fila(s) repetida(s))` : '';
      const sufijoFin = nextAlertas.length === 0 ? ' Todas las alertas resueltas.' : '';
      onNotify(`${prodFinal.nombre} creado y sumado${sufijoAuto}.${sufijoFin}`, 'green');
    } else {
      onNotify(
        fpGuardarEnCatalogo
          ? `${prodFinal.nombre} guardado en el catálogo y agregado al pedido.`
          : `${prodFinal.nombre} agregado solo a este pedido.`,
        'green'
      );
    }
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
              onClick={() => {
                if (mostrarFueraPlanilla) {
                  resetFueraPlanilla();
                  setMostrarFueraPlanilla(false);
                } else {
                  resetFueraPlanilla();
                  setMostrarFueraPlanilla(true);
                }
              }}
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
                    limpiarEstadoAlertas();
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
            ref={fueraPlanillaRef}
            onSubmit={(e) => void handleAgregarFueraPlanilla(e)}
            className="bg-[#FAFAFA] dark:bg-[#1B1B1F] rounded-2xl border border-zinc-200/80 dark:border-[#26262B] p-4 space-y-3.5 text-xs"
          >
            <div className="flex items-center justify-between">
              <div>
                <div className="font-semibold text-[#111111] dark:text-[#F4F4F5]">
                  {alertaCreandoId ? 'Crear producto desde alerta del Excel' : 'Agregar producto fuera de planilla'}
                </div>
                <div className="text-[11px] text-zinc-500 dark:text-zinc-400">
                  {alertaCreandoId
                    ? 'Confirmá nombre y precio real del catálogo antes de sumar al pedido.'
                    : 'Ej. Pilusos o artículos nuevos'}
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  resetFueraPlanilla();
                  setMostrarFueraPlanilla(false);
                }}
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
                <div className="text-[11px] text-zinc-400 dark:text-zinc-500 mt-1 tabular-nums flex flex-wrap items-center justify-between gap-1.5">
                  <span>Con IVA 21%: {formatARS(fpPrecioConIva)}</span>
                  {fpPrecioExcelHint !== null && fpPrecioExcelHint > 0 && (
                    <span className="inline-flex items-center gap-1.5 text-zinc-500 dark:text-zinc-400">
                      <span>En el Excel figura {formatARS(fpPrecioExcelHint)}</span>
                      <button
                        type="button"
                        onClick={() => setFpPrecioNeto(String(fpPrecioExcelHint))}
                        className="px-2 py-0.5 rounded-full bg-zinc-200/80 hover:bg-zinc-300 text-[#111111] dark:bg-[#2A2A32] dark:hover:bg-[#35353F] dark:text-[#F4F4F5] font-medium transition-colors"
                      >
                        Usar {formatARSCompacto(fpPrecioExcelHint)}
                      </button>
                    </span>
                  )}
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
                onClick={() => {
                  resetFueraPlanilla();
                  setMostrarFueraPlanilla(false);
                }}
                className="px-3.5 py-2 rounded-full border border-zinc-200 dark:border-[#2E2E35] bg-white dark:bg-[#141417] text-zinc-600 dark:text-zinc-300"
              >
                Cancelar
              </button>
              <button
                type="submit"
                disabled={fpGuardando}
                className="px-4 py-2 rounded-full bg-zinc-900 text-white font-medium hover:bg-zinc-800 dark:bg-zinc-100 dark:text-zinc-950 dark:hover:bg-white disabled:opacity-50"
              >
                {fpGuardando
                  ? 'Guardando...'
                  : alertaCreandoId
                    ? 'Crear y sumar al pedido'
                    : 'Agregar al pedido'}
              </button>
            </div>
          </form>
        )}
      </section>

      {(advertenciasExcel.length > 0 || alertasIgnoradas.length > 0) && (
        <section className="bg-red-50/70 border border-red-200 dark:bg-red-500/10 dark:border-red-500/25 rounded-2xl p-5 space-y-3.5">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-xs font-semibold text-red-800 dark:text-red-300">
              <AlertTriangle className="w-4 h-4 text-red-600 dark:text-red-400 shrink-0" strokeWidth={1.75} />
              <span>
                {advertenciasExcel.length > 0
                  ? `Alertas del Excel para resolver (${advertenciasExcel.length})`
                  : `${alertasIgnoradas.length} ${alertasIgnoradas.length === 1 ? 'fila ignorada' : 'filas ignoradas'} del Excel`}
              </span>
            </div>

            <div className="flex flex-wrap items-center gap-2 text-[11px]">
              {alertasIgnoradas.length > 0 && (
                <button
                  type="button"
                  onClick={handleRestaurarAlertasIgnoradas}
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-white hover:bg-zinc-50 border border-red-200/80 text-red-700 dark:bg-[#1B1B1F] dark:hover:bg-[#24242B] dark:border-red-500/30 dark:text-red-300 font-medium transition-colors"
                >
                  <RotateCcw className="w-3 h-3" strokeWidth={1.75} />
                  Restaurar ({alertasIgnoradas.length})
                </button>
              )}

              {advertenciasExcel.length > 0 &&
                (confirmIgnorarTodas ? (
                  <div className="inline-flex items-center gap-1.5 bg-red-100/90 border border-red-300 dark:bg-red-950/60 dark:border-red-500/40 rounded-full px-2.5 py-1">
                    <span className="text-red-900 dark:text-red-200 font-medium">
                      ¿Ignorar las {advertenciasExcel.length}?
                    </span>
                    <button
                      type="button"
                      onClick={handleIgnorarTodasAlertas}
                      className="px-2 py-0.5 rounded-full bg-red-600 text-white font-semibold"
                    >
                      Sí
                    </button>
                    <button
                      type="button"
                      onClick={() => setConfirmIgnorarTodas(false)}
                      className="px-1.5 py-0.5 text-zinc-700 dark:text-zinc-300"
                    >
                      No
                    </button>
                  </div>
                ) : (
                  <button
                    type="button"
                    onClick={() => setConfirmIgnorarTodas(true)}
                    className="font-medium text-red-700 dark:text-red-300 hover:underline"
                  >
                    Ignorar todas
                  </button>
                ))}
            </div>
          </div>

          {advertenciasExcel.length > 0 && (
            <div className="space-y-2.5">
              {advertenciasExcel.map((adv) => {
                const estaCorrigiendo = corrigiendoAlertaId === adv.id;
                const qCorr = normalizarTexto(busquedaCorreccion);
                const matchesCatalogo = estaCorrigiendo
                  ? catalogoCompleto
                      .filter((p) => {
                        if (!qCorr) return true;
                        return (
                          normalizarTexto(p.nombre).includes(qCorr) ||
                          normalizarTexto(p.rubro).includes(qCorr) ||
                          p.variantes.some((v) => normalizarTexto(v.codigo).includes(qCorr))
                        );
                      })
                      .slice(0, 6)
                  : [];

                return (
                  <div
                    key={adv.id}
                    className="rounded-xl bg-white dark:bg-[#1B1B1F] border border-red-200/80 dark:border-red-500/25 p-3.5 space-y-2.5 text-xs"
                  >
                    <div className="flex flex-wrap items-start justify-between gap-2">
                      <div className="space-y-1 min-w-0">
                        <div className="flex flex-wrap items-center gap-1.5 text-[11px] tabular-nums">
                          <span className="font-semibold text-[#111111] dark:text-[#F4F4F5]">
                            Hoja {adv.hoja} · Fila {adv.fila}
                          </span>
                          <span className="px-2 py-0.5 rounded-md bg-zinc-100 text-zinc-700 dark:bg-[#2A2A32] dark:text-zinc-200 font-medium uppercase">
                            Cód. {adv.codigo || '—'}
                          </span>
                          {adv.cantidadDetectada !== null && (
                            <span className="px-2 py-0.5 rounded-md bg-[#FCE7F3] text-[#111111] dark:bg-[#F4B6CE]/15 dark:text-[#FCE7F3] font-semibold">
                              {adv.cantidadDetectada} u.
                            </span>
                          )}
                          {adv.precioExcelReferencia !== null && (
                            <span className="px-2 py-0.5 rounded-md bg-zinc-50 text-zinc-500 border border-zinc-200/70 dark:bg-[#141417] dark:text-zinc-400 dark:border-[#2E2E35]">
                              En planilla: {formatARS(adv.precioExcelReferencia)}
                            </span>
                          )}
                        </div>

                        <div className="font-medium text-[#111111] dark:text-[#F4F4F5]">
                          {adv.descripcionExcel || adv.mensaje}
                        </div>

                        <div className="flex flex-wrap items-center gap-x-3 gap-y-0.5 text-[11px]">
                          <span className="text-red-600 dark:text-red-400">
                            {adv.tipo === 'CODIGO_DESCONOCIDO'
                              ? 'Código no reconocido en el catálogo vigente'
                              : adv.mensaje}
                          </span>
                          {adv.notaDetectada && (
                            <span className="text-zinc-500 dark:text-zinc-400">
                              Detalle: {adv.notaDetectada}
                            </span>
                          )}
                        </div>
                      </div>

                      {adv.tipo === 'CODIGO_DESCONOCIDO' && (
                        <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                          <button
                            type="button"
                            onClick={() => handleIniciarCrearDesdeAlerta(adv)}
                            className="inline-flex items-center gap-1 px-3 py-1.5 rounded-full text-[11px] font-medium bg-zinc-900 hover:bg-zinc-800 text-white dark:bg-zinc-100 dark:hover:bg-white dark:text-zinc-950 transition-colors"
                          >
                            <Plus className="w-3 h-3" strokeWidth={1.75} />
                            Crear y sumar
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setCorrigiendoAlertaId(estaCorrigiendo ? null : adv.id);
                              setBusquedaCorreccion('');
                            }}
                            className={`inline-flex items-center gap-1 px-3 py-1.5 rounded-full text-[11px] font-medium border transition-colors ${
                              estaCorrigiendo
                                ? 'bg-zinc-100 border-zinc-300 text-[#111111] dark:bg-[#2A2A32] dark:border-[#3A3A44] dark:text-white'
                                : 'bg-white hover:bg-zinc-50 border-zinc-200 text-zinc-700 dark:bg-[#141417] dark:hover:bg-[#24242B] dark:border-[#2E2E35] dark:text-zinc-200'
                            }`}
                          >
                            <Search className="w-3 h-3" strokeWidth={1.75} />
                            Corregir código
                          </button>
                          <button
                            type="button"
                            onClick={() => handleIgnorarAlerta(adv)}
                            className="px-2.5 py-1.5 rounded-full text-[11px] text-zinc-500 hover:text-red-700 hover:bg-red-50 dark:text-zinc-400 dark:hover:text-red-300 dark:hover:bg-red-500/10 transition-colors"
                          >
                            Ignorar fila
                          </button>
                        </div>
                      )}
                    </div>

                    {adv.tipo === 'CANTIDAD_INVALIDA' && (
                      <div className="flex flex-wrap items-center gap-2 pt-1 border-t border-zinc-100 dark:border-[#26262B]">
                        <span className="text-[11px] text-zinc-500 dark:text-zinc-400">
                          Ingresá cantidad correcta{adv.cantidadRaw ? ` (leído: "${adv.cantidadRaw}")` : ''}:
                        </span>
                        <input
                          type="number"
                          min={1}
                          step={1}
                          value={cantidadesCorreccion[adv.id] ?? ''}
                          onWheel={(e) => e.currentTarget.blur()}
                          onChange={(e) =>
                            setCantidadesCorreccion((prev) => ({ ...prev, [adv.id]: e.target.value }))
                          }
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              handleCorregirCantidadAlerta(adv);
                            }
                          }}
                          placeholder="Ej. 4"
                          className="w-24 px-2.5 py-1.5 text-xs bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg tabular-nums placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
                        />
                        <button
                          type="button"
                          onClick={() => handleCorregirCantidadAlerta(adv)}
                          className="px-3 py-1.5 rounded-full text-[11px] font-medium bg-zinc-900 hover:bg-zinc-800 text-white dark:bg-zinc-100 dark:hover:bg-white dark:text-zinc-950 transition-colors"
                        >
                          Aplicar
                        </button>
                        <button
                          type="button"
                          onClick={() => handleIgnorarAlerta(adv)}
                          className="px-2.5 py-1.5 rounded-full text-[11px] text-zinc-500 hover:text-red-700 hover:bg-red-50 dark:text-zinc-400 dark:hover:text-red-300 dark:hover:bg-red-500/10 transition-colors ml-auto"
                        >
                          Ignorar fila
                        </button>
                      </div>
                    )}

                    {adv.tipo === 'CODIGO_DESCONOCIDO' && estaCorrigiendo && (
                      <div className="pt-2.5 border-t border-zinc-200/70 dark:border-[#26262B] space-y-2.5">
                        {adv.sugerenciasCatalogo.length > 0 && (
                          <div className="space-y-1.5">
                            <div className="text-[11px] font-medium text-zinc-500 dark:text-zinc-400">
                              Coincidencias sugeridas del catálogo
                            </div>
                            <div className="flex flex-wrap gap-1.5">
                              {adv.sugerenciasCatalogo.map((sug) => {
                                const prodSug = catalogoCompleto.find((p) => p.idGrupo === sug.idGrupo);
                                if (!prodSug) return null;
                                return (
                                  <button
                                    key={sug.idGrupo}
                                    type="button"
                                    onClick={() => handleVincularAlertaConProducto(adv, prodSug)}
                                    className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-xl bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 text-left dark:bg-[#141417] dark:hover:bg-[#24242B] dark:border-[#2E2E35] transition-colors"
                                  >
                                    <span className="px-1.5 py-0.5 rounded bg-[#FCE7F3] text-[#111111] dark:bg-[#F4B6CE]/15 dark:text-[#FCE7F3] text-[10px] font-medium">
                                      {sug.motivo}
                                    </span>
                                    <span className="font-medium text-[#111111] dark:text-[#F4F4F5]">
                                      {sug.nombre}
                                    </span>
                                    <span className="text-[10px] text-zinc-400 dark:text-zinc-500 tabular-nums">
                                      ({sug.codigos.join('/')})
                                    </span>
                                  </button>
                                );
                              })}
                            </div>
                          </div>
                        )}

                        <div className="space-y-1.5">
                          <input
                            type="text"
                            value={busquedaCorreccion}
                            onChange={(e) => setBusquedaCorreccion(e.target.value)}
                            placeholder="Buscar producto del catálogo por nombre, código o rubro..."
                            className="w-full px-3 py-1.5 text-xs bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus:outline-none focus:border-zinc-900 dark:focus:border-[#F4B6CE]/70"
                          />
                          <div className="max-h-40 overflow-y-auto divide-y divide-zinc-100 dark:divide-[#26262B] border border-zinc-200/80 dark:border-[#26262B] rounded-xl">
                            {matchesCatalogo.length === 0 ? (
                              <div className="px-3 py-2 text-[11px] text-zinc-400 dark:text-zinc-500">
                                No hay coincidencias para esa búsqueda.
                              </div>
                            ) : (
                              matchesCatalogo.map((prod) => (
                                <button
                                  key={prod.idGrupo}
                                  type="button"
                                  onClick={() => handleVincularAlertaConProducto(adv, prod)}
                                  className="w-full px-3 py-2 flex items-center justify-between gap-2 text-left hover:bg-zinc-50 dark:hover:bg-[#24242B] transition-colors"
                                >
                                  <div className="min-w-0">
                                    <div className="font-medium text-[#111111] dark:text-[#F4F4F5] truncate">
                                      {prod.nombre}
                                    </div>
                                    <div className="text-[10px] text-zinc-400 dark:text-zinc-500 tabular-nums">
                                      {prod.variantes.map((v) => v.codigo).join(' / ')} · {prod.rubro}
                                    </div>
                                  </div>
                                  <span className="text-[11px] font-medium text-zinc-600 dark:text-zinc-300 shrink-0 tabular-nums">
                                    {formatARSCompacto(prod.precioNeto)}
                                  </span>
                                </button>
                              ))
                            )}
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </section>
      )}

      {avisosPrecio.length > 0 && (
        <section className="bg-amber-50/60 border border-amber-200/80 dark:bg-amber-500/10 dark:border-amber-500/25 rounded-2xl p-4 space-y-2.5">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-xs font-medium text-amber-900 dark:text-amber-200">
              <Info className="w-4 h-4 text-amber-600 dark:text-amber-400 shrink-0" strokeWidth={1.75} />
              <span>
                Se respetaron los precios netos vigentes del catálogo en {avisosPrecio.length}{' '}
                {avisosPrecio.length === 1 ? 'producto' : 'productos'} con importe distinto en el Excel.
              </span>
            </div>
            <button
              type="button"
              onClick={() => setMostrarDetallePrecios((prev) => !prev)}
              className="inline-flex items-center gap-1 text-[11px] font-medium text-amber-800 dark:text-amber-300 hover:underline shrink-0"
            >
              {mostrarDetallePrecios ? (
                <>
                  Ocultar detalle <ChevronUp className="w-3.5 h-3.5" strokeWidth={1.75} />
                </>
              ) : (
                <>
                  Ver detalle <ChevronDown className="w-3.5 h-3.5" strokeWidth={1.75} />
                </>
              )}
            </button>
          </div>

          {mostrarDetallePrecios && (
            <div className="overflow-x-auto rounded-xl border border-amber-200/70 dark:border-amber-500/20 bg-white dark:bg-[#1B1B1F]">
              <table className="w-full text-[11px] tabular-nums">
                <thead>
                  <tr className="border-b border-zinc-100 dark:border-[#26262B] text-zinc-500 dark:text-zinc-400">
                    <th className="text-left font-medium px-3 py-2">Código</th>
                    <th className="text-left font-medium px-3 py-2">Producto</th>
                    <th className="text-right font-medium px-3 py-2">Excel</th>
                    <th className="text-right font-medium px-3 py-2">Vigente aplicado</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-100 dark:divide-[#26262B]">
                  {avisosPrecio.map((ap) => (
                    <tr key={ap.codigo}>
                      <td className="px-3 py-1.5 font-medium text-zinc-600 dark:text-zinc-300">{ap.codigo}</td>
                      <td className="px-3 py-1.5 text-[#111111] dark:text-[#F4F4F5]">{ap.nombreProducto}</td>
                      <td className="px-3 py-1.5 text-right text-zinc-400 dark:text-zinc-500 line-through">
                        {formatARS(ap.precioPlanilla)}
                      </td>
                      <td className="px-3 py-1.5 text-right font-semibold text-emerald-700 dark:text-emerald-400">
                        {formatARS(ap.precioVigente)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
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

              <div className="max-h-96 overflow-y-auto divide-y divide-zinc-200/60 dark:divide-[#26262B]/80 border border-zinc-200/70 dark:border-[#26262B] rounded-xl">
                {lineasBorrador.map((l) => {
                  const prod = catalogoCompleto.find((p) => p.idGrupo === l.idGrupo);
                  const variantes = prod ? prod.variantes : l.desglose;
                  const esMultiResuelto = !l.requiereRevision && l.estado === 'OK' && variantes.length >= 2;
                  return (
                    <div key={l.idLinea} className="px-3.5 py-2.5 space-y-2 text-xs">
                      <div className="flex items-center justify-between gap-2.5">
                        <div className="min-w-0">
                          <div className="font-medium text-[#111111] dark:text-[#F4F4F5] truncate">
                            {l.nombreProducto}
                          </div>
                          <div className="text-[11px] text-zinc-400 dark:text-zinc-500 tabular-nums">
                            {l.requiereRevision
                              ? `Pendiente definir (${l.cantidadTotal} u.)`
                              : `${formatARSCompacto(l.precioUnitarioNeto)} c/u neto · ${l.cantidadTotal} u.`}
                          </div>
                        </div>
                        <div className="flex items-center gap-2 shrink-0 tabular-nums">
                          {!esMultiResuelto && (
                            <div className="inline-flex items-center rounded-lg border border-zinc-200 dark:border-[#2E2E35] bg-[#FAFAFA] dark:bg-[#1B1B1F]">
                              <button
                                type="button"
                                disabled={l.cantidadTotal <= 1}
                                onClick={() => handleAjusteRapidoLineaPaso3(l.idGrupo, l.cantidadTotal - 1)}
                                aria-label={`Restar ${l.nombreProducto}`}
                                className="p-1 text-zinc-600 hover:text-[#111111] dark:text-zinc-300 dark:hover:text-white disabled:opacity-35 disabled:cursor-not-allowed"
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
                                    handleAjusteRapidoLineaPaso3(l.idGrupo, n);
                                  }
                                }}
                                className="w-10 text-center text-xs font-medium bg-transparent text-[#111111] dark:text-[#F4F4F5] tabular-nums focus:outline-none"
                              />
                              <button
                                type="button"
                                onClick={() => handleAjusteRapidoLineaPaso3(l.idGrupo, l.cantidadTotal + 1)}
                                aria-label={`Sumar ${l.nombreProducto}`}
                                className="p-1 text-zinc-600 hover:text-[#111111] dark:text-zinc-300 dark:hover:text-white"
                              >
                                <Plus className="w-3 h-3" strokeWidth={1.75} />
                              </button>
                            </div>
                          )}
                          {esMultiResuelto && editandoLineaId !== l.idLinea && (
                            <button
                              type="button"
                              onClick={() => setEditandoLineaId(l.idLinea)}
                              className="text-[11px] text-zinc-400 hover:text-[#111111] dark:text-zinc-500 dark:hover:text-[#F4F4F5] inline-flex items-center gap-1"
                            >
                              <Edit3 className="w-3 h-3" strokeWidth={1.5} />
                              Repartir
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

                      {esMultiResuelto && editandoLineaId !== l.idLinea && (
                        <div className="flex flex-wrap gap-1.5">
                          {variantes.map((v) => {
                            const qtyVar = l.desglose.find((d) => d.codigo === v.codigo)?.cantidad ?? 0;
                            const esUltimaUnidad = qtyVar <= 1 && l.cantidadTotal <= 1;
                            return (
                              <div
                                key={v.codigo}
                                className="inline-flex items-center gap-1.5 rounded-lg bg-[#FAFAFA] dark:bg-[#1B1B1F] border border-zinc-200/80 dark:border-[#26262B] px-2 py-1 text-[11px]"
                              >
                                <span className="text-zinc-600 dark:text-zinc-300 truncate max-w-[140px]">
                                  {v.codigo} {v.etiqueta}
                                </span>
                                <div className="inline-flex items-center rounded border border-zinc-200 dark:border-[#2E2E35] bg-white dark:bg-[#141417]">
                                  <button
                                    type="button"
                                    disabled={qtyVar <= 0 || esUltimaUnidad}
                                    onClick={() => handleAjusteVariantePaso3(l.idGrupo, v.codigo, qtyVar - 1)}
                                    aria-label={`Restar ${v.etiqueta}`}
                                    className="p-0.5 text-zinc-600 hover:text-[#111111] dark:text-zinc-300 dark:hover:text-white disabled:opacity-30 disabled:cursor-not-allowed"
                                  >
                                    <Minus className="w-2.5 h-2.5" strokeWidth={1.75} />
                                  </button>
                                  <input
                                    type="number"
                                    min={0}
                                    value={qtyVar}
                                    onWheel={(e) => e.currentTarget.blur()}
                                    onChange={(e) => {
                                      const n = parseInt(e.target.value, 10);
                                      handleAjusteVariantePaso3(
                                        l.idGrupo,
                                        v.codigo,
                                        Number.isFinite(n) ? Math.max(0, n) : 0
                                      );
                                    }}
                                    className="w-7 text-center text-[11px] font-medium bg-transparent text-[#111111] dark:text-[#F4F4F5] tabular-nums focus:outline-none"
                                  />
                                  <button
                                    type="button"
                                    onClick={() => handleAjusteVariantePaso3(l.idGrupo, v.codigo, qtyVar + 1)}
                                    aria-label={`Sumar ${v.etiqueta}`}
                                    className="p-0.5 text-zinc-600 hover:text-[#111111] dark:text-zinc-300 dark:hover:text-white"
                                  >
                                    <Plus className="w-2.5 h-2.5" strokeWidth={1.75} />
                                  </button>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}

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
