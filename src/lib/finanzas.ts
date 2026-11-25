import type {
  AlertaCheque,
  EstadoArmado,
  EstadoPago,
  LineaPedido,
  Pedido,
  PedidoBase,
  ResumenFinanciero
} from './types';

/** Convierte un monto en pesos a centavos enteros exactos (cero aritmética float acumulativa). */
export const toCents = (pesos: number): number => {
  if (!Number.isFinite(pesos)) return 0;
  return Math.round((pesos + Math.sign(pesos) * 1e-8) * 100);
};

/** Convierte centavos enteros a pesos con 2 decimales exactos. */
export const fromCents = (cents: number): number => {
  if (!Number.isFinite(cents)) return 0;
  return Math.trunc(Math.round(cents)) / 100;
};

/** Redondeo financiero a 2 decimales pasando por centavos enteros. */
export const round2 = (val: number): number => fromCents(toCents(val));

/** Aplica un porcentaje entero/decimal en centésimas sobre centavos enteros con redondeo half-up. */
const multiplyPercentCents = (cents: number, ratePercent: number): number => {
  if (!Number.isFinite(cents) || !Number.isFinite(ratePercent)) return 0;
  const scaled = Math.abs(cents) * Math.abs(ratePercent);
  const rounded = Math.floor((scaled + 50) / 100);
  return (cents < 0 ? -1 : 1) * rounded;
};

export const getTodayISO = (): string => {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

export const addDaysISO = (isoDate: string, days: number): string => {
  const base = /^\d{4}-\d{2}-\d{2}$/.test(isoDate) ? isoDate : getTodayISO();
  const dt = new Date(`${base}T12:00:00`);
  dt.setDate(dt.getDate() + days);
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  const d = String(dt.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

export const formatShortDate = (isoDate: string): string => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return isoDate || '—';
  const parts = isoDate.split('-');
  const m = parts[1] ?? '';
  const d = parts[2] ?? '';
  return `${d}/${m}`;
};

export const formatNowDateTime = (): string => {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  return `${dd}/${mm} ${hh}:${min}`;
};

export const formatARS = (val: number): string =>
  new Intl.NumberFormat('es-AR', {
    style: 'currency',
    currency: 'ARS',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(round2(val));

export const formatARSCompacto = (val: number): string => {
  const r = round2(val);
  const sinCentavos = Math.abs(r - Math.round(r)) < 0.005;
  return new Intl.NumberFormat('es-AR', {
    style: 'currency',
    currency: 'ARS',
    minimumFractionDigits: sinCentavos ? 0 : 2,
    maximumFractionDigits: sinCentavos ? 0 : 2
  }).format(r);
};

export const getChequeAlertStatus = (pedido: Pedido): AlertaCheque => {
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

export const getDiasParaAcreditacionCheque = (pedido: Pedido): number | null => {
  if (pedido.medioPago !== 'Cheque' || pedido.chequeAcreditado || pedido.montoCheque <= 0) {
    return null;
  }
  const venc = pedido.fechaAcreditacionCheque;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(venc)) return null;
  const hoyMs = new Date(`${getTodayISO()}T12:00:00`).getTime();
  const vencMs = new Date(`${venc}T12:00:00`).getTime();
  return Math.round((vencMs - hoyMs) / 86400000);
};

export const esDespachoAtrasado = (pedido: Pedido): boolean => {
  if (pedido.estadoEnvio !== 'Pendiente') return false;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(pedido.despachoPrevisto)) return false;
  return pedido.despachoPrevisto < getTodayISO();
};

export const construirDetalleResumen = (lineas: ReadonlyArray<LineaPedido>, fallback?: string): string => {
  if (lineas.length === 0) return fallback?.trim() || '—';
  const rubros = Array.from(new Set(lineas.map((l) => l.rubro)));
  const totalUnidades = lineas.reduce((acc, l) => acc + l.cantidadTotal, 0);
  return `${rubros.join(' · ')} (${totalUnidades} u.)`;
};

export const recalcularPedido = (base: PedidoBase): Pedido => {
  const descuentoPorc = Math.max(0, Math.min(100, Number.isFinite(base.descuentoPorc) ? round2(base.descuentoPorc) : 0));

  let subtotalCents = 0;
  for (const linea of base.lineas) {
    const unitCents = toCents(linea.precioUnitarioNeto);
    subtotalCents += unitCents * Math.max(0, Math.trunc(linea.cantidadTotal));
  }

  const descuentoCents = multiplyPercentCents(subtotalCents, descuentoPorc);
  const baseImponibleCents = subtotalCents - descuentoCents;
  const iva21Cents = multiplyPercentCents(baseImponibleCents, 21);
  const totalFacturadoCents = baseImponibleCents + iva21Cents;

  const usarOverride =
    base.lineas.length === 0 &&
    typeof base.totalFacturadoOverride === 'number' &&
    Number.isFinite(base.totalFacturadoOverride) &&
    base.totalFacturadoOverride >= 0;

  let finalSubtotalCents = subtotalCents;
  let finalDescuentoCents = descuentoCents;
  let finalBaseCents = baseImponibleCents;
  let finalIvaCents = iva21Cents;
  let finalTotalCents = totalFacturadoCents;

  if (usarOverride) {
    finalTotalCents = Math.max(0, toCents(base.totalFacturadoOverride ?? 0));
    finalBaseCents = Math.round((finalTotalCents * 100) / 121);
    finalIvaCents = finalTotalCents - finalBaseCents;
    finalSubtotalCents = finalBaseCents;
    finalDescuentoCents = 0;
  }

  let abonadoCents = Math.max(0, toCents(base.montoAbonado));
  if (abonadoCents > finalTotalCents) {
    abonadoCents = finalTotalCents;
  }

  let chequeCents = 0;
  if (base.medioPago === 'Cheque' && !base.chequeAcreditado) {
    const inputChequeCents = toCents(base.montoCheque);
    chequeCents = Math.max(0, inputChequeCents > 0 ? inputChequeCents : finalTotalCents - abonadoCents);
  }
  if (abonadoCents + chequeCents > finalTotalCents) {
    chequeCents = Math.max(0, finalTotalCents - abonadoCents);
  }

  const saldoCents = Math.max(0, finalTotalCents - abonadoCents);
  const deudaSinRespaldoCents = Math.max(0, finalTotalCents - abonadoCents - chequeCents);

  let estadoPago: EstadoPago;
  if (base.medioPago === 'Cheque' && !base.chequeAcreditado) {
    estadoPago = saldoCents <= 1 ? 'Pagado' : 'Pendiente';
  } else if (saldoCents <= 1 && finalTotalCents > 0) {
    estadoPago = 'Pagado';
  } else if (abonadoCents > 0) {
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
    montoAbonado: fromCents(abonadoCents),
    montoCheque: fromCents(chequeCents),
    estadoArmado,
    subtotalNeto: fromCents(finalSubtotalCents),
    montoDescuento: fromCents(finalDescuentoCents),
    baseImponible: fromCents(finalBaseCents),
    iva21: fromCents(finalIvaCents),
    totalFacturado: fromCents(finalTotalCents),
    saldoPendiente: fromCents(saldoCents),
    deudaSinRespaldo: fromCents(deudaSinRespaldoCents),
    estadoPago,
    faltantesVariante,
    detalleResumen: construirDetalleResumen(base.lineas, base.detalleManualImportado),
    incidenciaDeposito: base.incidenciaDeposito ?? '',
    incidenciaAt: base.incidenciaAt ?? null,
    incidenciaPor: base.incidenciaPor ?? '',
    asignaciones: Array.isArray(base.asignaciones) ? base.asignaciones : []
  };
};

export const calcularResumenFinanciero = (pedidos: ReadonlyArray<Pedido>): ResumenFinanciero => {
  let porArmar = 0;
  let cobradoEnManoCents = 0;
  let enChequesCents = 0;
  let chequesPorVencer = 0;
  let faltaPagoCents = 0;

  for (const p of pedidos) {
    if (
      p.estadoArmado === 'Por armar' ||
      p.estadoArmado === 'En armado' ||
      p.estadoArmado === 'Falta talle o color'
    ) {
      porArmar += 1;
    }
    cobradoEnManoCents += toCents(p.montoAbonado);
    if (p.medioPago === 'Cheque' && !p.chequeAcreditado && p.montoCheque > 0) {
      enChequesCents += toCents(p.montoCheque);
      const alerta = getChequeAlertStatus(p);
      if (alerta === 'VENCIDO' || alerta === 'VENCE_HOY') {
        chequesPorVencer += 1;
      }
    }
    faltaPagoCents += toCents(p.deudaSinRespaldo);
  }

  return {
    porArmar,
    cobradoEnMano: fromCents(cobradoEnManoCents),
    enCheques: fromCents(enChequesCents),
    chequesPorVencer,
    faltaPago: fromCents(faltaPagoCents)
  };
};
