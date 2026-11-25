import { CATALOGO_BY_CODIGO } from './catalogo-mibi';
import {
  aplicarResolucionLinea,
  crearLineasEjemploLista006,
  limpiarTextoInvisible,
  parsearCantidadEstricta,
  parsearMontoArgentino,
  resolverLinea
} from './excel-parser';
import {
  addDaysISO,
  formatARS,
  getChequeAlertStatus,
  getDiasParaAcreditacionCheque,
  getTodayISO,
  recalcularPedido,
  round2
} from './finanzas';
import { generarMensajeDespachoWhatsApp, generarMensajeWhatsApp } from './whatsapp';
import type { PedidoBase, ProductoCatalogo, ResultadoQA } from './types';

export const ejecutarSuiteQA = (): ResultadoQA[] => {
  const resultados: ResultadoQA[] = [];
  const check = (nombre: string, ok: boolean, detalle: string): void => {
    resultados.push({ nombre, ok, detalle });
  };

  const monster = CATALOGO_BY_CODIGO.get('6502')!;
  const shark = CATALOGO_BY_CODIGO.get('6504')!;
  const zapa = CATALOGO_BY_CODIGO.get('6576')!;
  const dino = CATALOGO_BY_CODIGO.get('6510')!;

  // 1-4. Redondeos exactos en centavos enteros y catálogo
  check('Redondeo 0.1 + 0.2 === 0.30', round2(0.1 + 0.2) === 0.3, `Resultado: ${round2(0.1 + 0.2)}`);
  check('Redondeo 35.855 === 35.86', round2(35.855) === 35.86, `Resultado: ${round2(35.855)}`);
  check('Redondeo 1.005 === 1.01', round2(1.005) === 1.01, `Resultado: ${round2(1.005)}`);
  check('Catálogo completo (6500-6614)', CATALOGO_BY_CODIGO.size >= 62, `${CATALOGO_BY_CODIGO.size} códigos activos`);

  // 5-14. Reglas estrictas de desglose
  const lUnica = resolverLinea(dino, 3, '');
  check(
    'Variante única (Antiparras Dino) → OK automático',
    lUnica.estado === 'OK' && (lUnica.desglose[0]?.cantidad ?? 0) === 3,
    lUnica.estado
  );

  const lSimetrica = resolverLinea(shark, 6, '3 y 3');
  check(
    'Nota simétrica "3 y 3" sobre 6 → OK (3 y 3)',
    lSimetrica.estado === 'OK' &&
      (lSimetrica.desglose[0]?.cantidad ?? 0) === 3 &&
      (lSimetrica.desglose[1]?.cantidad ?? 0) === 3,
    lSimetrica.estado
  );

  const lAsimetrica = resolverLinea(shark, 3, '2 y 1');
  check('Nota asimétrica "2 y 1" → REQUIERE_REVISION', lAsimetrica.requiereRevision, lAsimetrica.motivoRevision ?? '');

  const lRango = resolverLinea(shark, 3, 'de 1 a 3');
  check(
    'Nota "de 1 a 3" sobre 3 → todo a T1 (6504)',
    lRango.estado === 'OK' && (lRango.desglose[0]?.cantidad ?? 0) === 3,
    lRango.desglose[0]?.codigo ?? ''
  );

  const lRangoMismatch = resolverLinea(shark, 4, 'de 1 a 3: 2');
  check(
    'Nota "de 1 a 3: 2" sobre 4 → REQUIERE_REVISION',
    lRangoMismatch.requiereRevision,
    lRangoMismatch.motivoRevision ?? ''
  );

  const lGris = resolverLinea(monster, 3, 'gris');
  check('Nota "gris" en Monster → REQUIERE_REVISION', lGris.requiereRevision, lGris.motivoRevision ?? '');

  const lSurtido = resolverLinea(monster, 4, 'surtido');
  check('Nota "surtido" → REQUIERE_REVISION', lSurtido.requiereRevision, lSurtido.motivoRevision ?? '');

  const lVacia = resolverLinea(monster, 3, '');
  check('Nota vacía en 2 variantes → REQUIERE_REVISION', lVacia.requiereRevision, lVacia.motivoRevision ?? '');

  const lImpar = resolverLinea(monster, 3, 'mitad y mitad');
  check('"mitad y mitad" impar (3) → REQUIERE_REVISION', lImpar.requiereRevision, lImpar.motivoRevision ?? '');

  const lFalsoPositivo = resolverLinea(monster, 2, 'mejor 2');
  check(
    'Falso positivo "mejor 2" no asigna color → REQUIERE_REVISION',
    lFalsoPositivo.requiereRevision,
    lFalsoPositivo.motivoRevision ?? ''
  );

  // 15-19. Finanzas y tope de monto abonado
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
    version: 1,
    deletedAt: null,
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
  check(
    'Tope obligatorio: montoAbonado <= totalFacturado',
    pedTope.montoAbonado === pedTope.totalFacturado && pedTope.saldoPendiente === 0,
    `${formatARS(pedTope.montoAbonado)}`
  );

  const pedCheque = recalcularPedido({
    ...pedBase,
    lineas: [lUnica],
    medioPago: 'Cheque',
    montoCheque: pedTope.totalFacturado,
    chequeAcreditado: false,
    fechaRecepcionCheque: getTodayISO(),
    fechaAcreditacionCheque: addDaysISO(getTodayISO(), 30)
  });
  check(
    'Cheque sin acreditar → estado Pendiente',
    pedCheque.estadoPago === 'Pendiente' && pedCheque.deudaSinRespaldo === 0,
    pedCheque.estadoPago
  );

  const pedChequeVencido = recalcularPedido({ ...pedCheque, fechaAcreditacionCheque: addDaysISO(getTodayISO(), -2) });
  check(
    'Cheque vencido → alerta VENCIDO',
    getChequeAlertStatus(pedChequeVencido) === 'VENCIDO',
    getChequeAlertStatus(pedChequeVencido) ?? ''
  );

  const pedChequeHoy = recalcularPedido({ ...pedCheque, fechaAcreditacionCheque: getTodayISO() });
  check(
    'Cheque con vencimiento hoy → 0 días (Al cobro)',
    getDiasParaAcreditacionCheque(pedChequeHoy) === 0 && getChequeAlertStatus(pedChequeHoy) === 'VENCE_HOY',
    `Días: ${getDiasParaAcreditacionCheque(pedChequeHoy)}`
  );

  // 20-26. Validaciones Parser
  const qInvalida = parsearCantidadEstricta('3 docenas');
  check('Rechaza cantidad "3 docenas"', !qInvalida.valid, qInvalida.motivo ?? '');

  const qNegativa = parsearCantidadEstricta('-2');
  check('Rechaza cantidad negativa "-2"', !qNegativa.valid, qNegativa.motivo ?? '');

  const qDecimal = parsearCantidadEstricta('1,5');
  check('Rechaza cantidad decimal "1,5"', !qDecimal.valid, qDecimal.motivo ?? '');

  const limpioInv = limpiarTextoInvisible('\u200B6504\uFEFF');
  check('Limpia caracteres invisibles \\u200B/\\uFEFF', limpioInv === '6504', limpioInv);

  const lExplicita = resolverLinea(zapa, 4, '2 t1 y 2 t2');
  check(
    'Desglose explícito "2 t1 y 2 t2" → OK',
    lExplicita.estado === 'OK' &&
      (lExplicita.desglose[0]?.cantidad ?? 0) === 2 &&
      (lExplicita.desglose[1]?.cantidad ?? 0) === 2,
    lExplicita.estado
  );

  const desc10 = recalcularPedido({ ...pedBase, lineas: [lUnica], descuentoPorc: 10 });
  const baseEsp = round2(round2(3 * 9505) * 0.9);
  const totEsp = round2(baseEsp * 1.21);
  check(
    'Descuento 10% + IVA 21% exacto',
    desc10.baseImponible === baseEsp && desc10.totalFacturado === totEsp,
    formatARS(desc10.totalFacturado)
  );

  check(
    'Demo Lista 006: 1 artículo frenado (Monster)',
    pedDemo.faltantesVariante === 1 && pedDemo.estadoArmado === 'Falta talle o color',
    `Faltantes: ${pedDemo.faltantesVariante}`
  );

  const msgDespacho = generarMensajeDespachoWhatsApp({
    ...pedCheque,
    cliente: 'Playa Kids',
    numeroPedido: '0042',
    transporte: 'Andreani',
    nroSeguimiento: 'AND-9981'
  });
  check(
    'Mensaje WhatsApp de envío con transporte y guía',
    msgDespacho.includes('Playa Kids') &&
      msgDespacho.includes('#0042') &&
      msgDespacho.includes('Andreani') &&
      msgDespacho.includes('AND-9981'),
    'Texto listo para copiar'
  );

  const montoFecha = parsearMontoArgentino('10/11');
  check('No confunde fecha "10/11" con monto', montoFecha === 0, `Resultado: ${montoFecha}`);

  const pilusoCat: ProductoCatalogo = {
    idGrupo: 'CUSTOM_PILUSO_SHARK',
    nombre: 'Piluso Shark',
    rubro: 'PILUSOS',
    precioNeto: 8500,
    esCustom: true,
    variantes: [
      { codigo: 'PIL-S', etiqueta: 'T1 (1-3 años)', aliases: ['PIL-S', 'T1'] },
      { codigo: 'PIL-M', etiqueta: 'T2 (3-6 años)', aliases: ['PIL-M', 'T2'] },
      { codigo: 'PIL-L', etiqueta: 'T3 (6-9 años)', aliases: ['PIL-L', 'T3'] }
    ]
  };
  const lPilPend = resolverLinea(pilusoCat, 6, 'surtido');
  const lPilRes = aplicarResolucionLinea(lPilPend, [2, 3, 1], [pilusoCat]);
  const msgPil = generarMensajeWhatsApp('Juguetería Sol', '0088', [lPilPend], [pilusoCat]);
  check(
    'Producto custom N variantes (3 talles): resolución y WhatsApp completo',
    lPilPend.requiereRevision === true &&
      lPilRes !== null &&
      lPilRes.requiereRevision === false &&
      lPilRes.desglose.length === 3 &&
      msgPil.includes('PIL-S') &&
      msgPil.includes('PIL-L'),
    lPilRes ? `${lPilRes.desglose.map((d) => `${d.codigo}:${d.cantidad}`).join(', ')}` : 'Error'
  );

  const pedCustomCalc = recalcularPedido({
    ...pedBase,
    lineas: lPilRes ? [lPilRes] : [],
    descuentoPorc: 10
  });
  const pilBase = round2(round2(6 * 8500) * 0.9);
  const pilTot = round2(pilBase * 1.21);
  check(
    'Finanzas centavos enteros en producto fuera de planilla',
    pedCustomCalc.baseImponible === pilBase && pedCustomCalc.totalFacturado === pilTot,
    formatARS(pedCustomCalc.totalFacturado)
  );

  return resultados;
};
