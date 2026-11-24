import { describe, expect, it, vi, beforeEach } from 'vitest';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { CATALOGO_BY_CODIGO, CATALOGO_MIBI } from '@/lib/catalogo-mibi';
import {
  addDaysISO,
  calcularResumenFinanciero,
  fromCents,
  getChequeAlertStatus,
  getDiasParaAcreditacionCheque,
  getTodayISO,
  recalcularPedido,
  round2,
  toCents
} from '@/lib/finanzas';
import {
  aplicarResolucionLinea,
  crearLineasEjemploLista006,
  importarPlanillaGeneral,
  limpiarTextoInvisible,
  parsearCantidadEstricta,
  parsearExcelLista006,
  parsearFechaExcel,
  parsearMontoArgentino,
  resolverLinea
} from '@/lib/excel-parser';
import {
  generarBufferHojaDeposito,
  generarBufferPickingConsolidado,
  generarBufferTableroMaestro
} from '@/lib/excel-exporter';
import { generarMensajeDespachoWhatsApp, generarMensajeWhatsApp } from '@/lib/whatsapp';
import { ejecutarSuiteQA } from '@/lib/qa-suite';
import { hashPassword, verifyPassword } from '@/lib/auth/password';
import type { PedidoBase, SessionUser } from '@/lib/types';

let mockSession: SessionUser = {
  userId: '00000000-0000-4000-8000-000000000001',
  username: 'admin',
  nombre: 'Administración',
  rol: 'ADMIN',
  tokenVersion: 1
};

vi.mock('@/lib/auth/session', () => ({
  getSessionFromCookies: async () => mockSession,
  setSessionCookie: async () => {},
  clearSessionCookie: async () => {}
}));

import {
  batchUpdatePedidosDal,
  crearPedidoDal,
  patchPedidoDal,
  reiniciarSistemaDal,
  restorePedidosDal,
  softDeletePedidosDal
} from '@/lib/dal/pedidos';
import { AuthError } from '@/lib/dal/auth';

const makePedidoBase = (overrides: Partial<PedidoBase> = {}): PedidoBase => ({
  id: 'ped_test_1',
  numeroPedido: '0001',
  cliente: 'Casa Lorenzo',
  vendedor: 'Juancho',
  whatsapp: '11-5606-9384',
  fechaPedido: getTodayISO(),
  despachoPrevisto: addDaysISO(getTodayISO(), 5),
  lineas: [],
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
  historial: [],
  ...overrides
});

describe('MIBI Core — Catálogo, Centavos, Excel, Seguridad y RBAC', () => {
  beforeEach(() => {
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
  });

  // 1-4: Catálogo y aritmética exacta en centavos
  it('1. Catálogo oficial contiene los 54 grupos y 69 códigos (6500-6614)', () => {
    expect(CATALOGO_MIBI.length).toBe(54);
    expect(CATALOGO_BY_CODIGO.size).toBe(69);
    expect(CATALOGO_BY_CODIGO.get('6502')?.precioNeto).toBe(42645);
    expect(CATALOGO_BY_CODIGO.get('6614')?.precioNeto).toBe(3500);
  });

  it('2. Conversión exacta pesos <-> centavos enteros sin deriva IEEE-754', () => {
    expect(toCents(0.1) + toCents(0.2)).toBe(30);
    expect(fromCents(toCents(0.1) + toCents(0.2))).toBe(0.3);
    expect(round2(35.855)).toBe(35.86);
    expect(round2(1.005)).toBe(1.01);
  });

  it('3. Caso real auditado Lista N° 006: $2.942.705,00 neto / $617.968,05 IVA / $3.560.673,05 total', () => {
    const p = recalcularPedido(makePedidoBase({ lineas: crearLineasEjemploLista006() }));
    expect(p.subtotalNeto).toBe(2942705);
    expect(p.iva21).toBe(617968.05);
    expect(p.totalFacturado).toBe(3560673.05);
    expect(p.faltantesVariante).toBe(1);
    expect(p.estadoArmado).toBe('Falta talle o color');
  });

  it('4. Descuento 10% + IVA 21% calculado exactamente en centavos', () => {
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const p = recalcularPedido(makePedidoBase({ lineas: [resolverLinea(dino, 3, '')], descuentoPorc: 10 }));
    expect(p.subtotalNeto).toBe(28515);
    expect(p.montoDescuento).toBe(2851.5);
    expect(p.baseImponible).toBe(25663.5);
    expect(p.iva21).toBe(5389.34);
    expect(p.totalFacturado).toBe(31052.84);
  });

  // 5-13: Reglas determinísticas de talles/colores
  it('5. Variante única pasa a OK automáticamente', () => {
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const l = resolverLinea(dino, 5, '');
    expect(l.estado).toBe('OK');
    expect(l.requiereRevision).toBe(false);
    expect(l.desglose[0]?.cantidad).toBe(5);
  });

  it('6. Nota simétrica "3 y 3" sobre 6 -> OK', () => {
    const shark = CATALOGO_BY_CODIGO.get('6504')!;
    const l = resolverLinea(shark, 6, '3 y 3');
    expect(l.estado).toBe('OK');
    expect(l.desglose[0]?.cantidad).toBe(3);
    expect(l.desglose[1]?.cantidad).toBe(3);
  });

  it('7. Nota asimétrica "2 y 1" -> REQUIERE_REVISION', () => {
    const shark = CATALOGO_BY_CODIGO.get('6504')!;
    const l = resolverLinea(shark, 3, '2 y 1');
    expect(l.estado).toBe('REQUIERE_REVISION');
    expect(l.requiereRevision).toBe(true);
  });

  it('8. Nota "de 1 a 3" sobre 3 -> todo a T1 (6504)', () => {
    const shark = CATALOGO_BY_CODIGO.get('6504')!;
    const l = resolverLinea(shark, 3, 'de 1 a 3');
    expect(l.estado).toBe('OK');
    expect(l.desglose[0]?.codigo).toBe('6504');
    expect(l.desglose[0]?.cantidad).toBe(3);
  });

  it('9. Nota "de 1 a 3: 2" sobre 4 -> REQUIERE_REVISION', () => {
    const shark = CATALOGO_BY_CODIGO.get('6504')!;
    const l = resolverLinea(shark, 4, 'de 1 a 3: 2');
    expect(l.requiereRevision).toBe(true);
  });

  it('10. Nota "surtido" o color inexistente "gris" -> REQUIERE_REVISION', () => {
    const monster = CATALOGO_BY_CODIGO.get('6502')!;
    expect(resolverLinea(monster, 4, 'surtido').requiereRevision).toBe(true);
    expect(resolverLinea(monster, 3, 'gris').requiereRevision).toBe(true);
    expect(resolverLinea(monster, 3, '').requiereRevision).toBe(true);
  });

  it('11. "mitad y mitad" impar -> REQUIERE_REVISION; par -> OK', () => {
    const monster = CATALOGO_BY_CODIGO.get('6502')!;
    expect(resolverLinea(monster, 3, 'mitad y mitad').requiereRevision).toBe(true);
    const ok = resolverLinea(monster, 4, 'mitad y mitad');
    expect(ok.estado).toBe('OK');
    expect(ok.desglose[0]?.cantidad).toBe(2);
    expect(ok.desglose[1]?.cantidad).toBe(2);
  });

  it('12. Previene falso positivo de alias de 1 letra ("mejor 2") y respeta "2 rosa 1 verde"', () => {
    const monster = CATALOGO_BY_CODIGO.get('6502')!;
    expect(resolverLinea(monster, 2, 'mejor 2').requiereRevision).toBe(true);
    const explicit = resolverLinea(monster, 3, '2 rosa 1 verde');
    expect(explicit.estado).toBe('OK');
    expect(explicit.desglose[0]?.cantidad).toBe(1); // Green
    expect(explicit.desglose[1]?.cantidad).toBe(2); // Pink
  });

  it('13. aplicarResolucionLinea valida suma exacta entera no negativa', () => {
    const monster = CATALOGO_BY_CODIGO.get('6502')!;
    const pendiente = resolverLinea(monster, 4, '');
    expect(aplicarResolucionLinea(pendiente, 2, 1)).toBeNull();
    expect(aplicarResolucionLinea(pendiente, -1, 5)).toBeNull();
    const ok = aplicarResolucionLinea(pendiente, 3, 1);
    expect(ok?.estado).toBe('OK');
    expect(ok?.desglose[0]?.cantidad).toBe(3);
  });

  // 14-18: Finanzas, topes y alertas de cheques
  it('14. Tope obligatorio montoAbonado <= totalFacturado', () => {
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const p = recalcularPedido(makePedidoBase({ lineas: [resolverLinea(dino, 2, '')], montoAbonado: 9999999 }));
    expect(p.montoAbonado).toBe(p.totalFacturado);
    expect(p.saldoPendiente).toBe(0);
    expect(p.estadoPago).toBe('Pagado');
  });

  it('15. Cheques: estados VENCIDO, VENCE_HOY (0 días) y EN_PLAZO', () => {
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const linea = resolverLinea(dino, 2, '');
    const hoy = getTodayISO();

    const pHoy = recalcularPedido(
      makePedidoBase({
        lineas: [linea],
        medioPago: 'Cheque',
        chequeAcreditado: false,
        fechaRecepcionCheque: addDaysISO(hoy, -30),
        fechaAcreditacionCheque: hoy
      })
    );
    expect(pHoy.estadoPago).toBe('Pendiente');
    expect(pHoy.deudaSinRespaldo).toBe(0);
    expect(getDiasParaAcreditacionCheque(pHoy)).toBe(0);
    expect(getChequeAlertStatus(pHoy)).toBe('VENCE_HOY');

    const pVencido = recalcularPedido({ ...pHoy, fechaAcreditacionCheque: addDaysISO(hoy, -2) });
    expect(getChequeAlertStatus(pVencido)).toBe('VENCIDO');

    const pPlazo = recalcularPedido({ ...pHoy, fechaAcreditacionCheque: addDaysISO(hoy, 10) });
    expect(getChequeAlertStatus(pPlazo)).toBe('EN_PLAZO');
  });

  it('16. Resumen financiero agrega correctamente los 4 indicadores', () => {
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const p1 = recalcularPedido(
      makePedidoBase({ id: '1', lineas: [resolverLinea(dino, 1, '')], medioPago: 'Transferencia / Efectivo', montoAbonado: 11501.05 })
    );
    const p2 = recalcularPedido(
      makePedidoBase({
        id: '2',
        lineas: [resolverLinea(dino, 1, '')],
        medioPago: 'Cheque',
        fechaAcreditacionCheque: getTodayISO()
      })
    );
    const p3 = recalcularPedido(makePedidoBase({ id: '3', lineas: [resolverLinea(dino, 1, '')], medioPago: 'Sin pagar' }));
    const res = calcularResumenFinanciero([p1, p2, p3]);
    expect(res.porArmar).toBe(3);
    expect(res.cobradoEnMano).toBe(11501.05);
    expect(res.enCheques).toBe(11501.05);
    expect(res.chequesPorVencer).toBe(1);
    expect(res.faltaPago).toBe(11501.05);
  });

  // 17-24: Parser Excel & CSV & Importador histórico
  it('17. parsearCantidadEstricta rechaza texto, negativos y decimales', () => {
    expect(parsearCantidadEstricta('3 docenas').valid).toBe(false);
    expect(parsearCantidadEstricta('-2').valid).toBe(false);
    expect(parsearCantidadEstricta('1,5').valid).toBe(false);
    expect(parsearCantidadEstricta('4')).toEqual({ valid: true, qty: 4 });
  });

  it('18. limpiarTextoInvisible elimina BOM y zero-width spaces', () => {
    expect(limpiarTextoInvisible('\u200B6504\uFEFF')).toBe('6504');
  });

  it('19. parsearExcelLista006 exige cabecera LISTA N° 006 y rechaza archivos ajenos', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['Otra planilla'], ['6510', '', '', 9505, 2, '', '']]);
    XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
    const res = parsearExcelLista006(buf, 'ajena.xlsx');
    expect(res.ok).toBe(false);
  });

  it('20. parsearExcelLista006 ignora fila duplicada 6545 vacía y detecta códigos desconocidos', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['MIBI LISTA N° 006'],
      ['Codigos', '', 'Producto', 'Precio', 'Cantidades pedidas', '', 'Talles/Colores'],
      ['6545', '', 'Crabby duplicada', '', '', '', ''],
      ['6510', '', 'Antiparras Dino', 9505, 3, '', ''],
      ['9999', '', 'Desconocido', 1000, 2, '', '']
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Lista');
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
    const res = parsearExcelLista006(buf, 'lista006.xlsx');
    expect(res.ok).toBe(true);
    expect(res.lineas).toHaveLength(1);
    expect(res.lineas[0]?.idGrupo).toBe('6510');
    expect(res.advertencias).toHaveLength(1);
    expect(res.advertencias[0]?.codigo).toBe('9999');
  });

  it('21. parsearExcelLista006 soporta CSV argentino (;) e ISO-8859-1 y CSV estándar (,)', () => {
    const csvSemi = 'MIBI LISTA N° 006;;;;;;\nCodigos;;Producto;Precio;Cantidades pedidas;;Talles/Colores\n6504;6505;Shark;46612;4;;2 y 2';
    const bufSemi = new TextEncoder().encode(csvSemi).buffer;
    const resSemi = parsearExcelLista006(bufSemi, 'pedido.csv');
    expect(resSemi.ok).toBe(true);
    expect(resSemi.lineas[0]?.estado).toBe('OK');

    const csvComma = 'MIBI LISTA N° 006,,,,,,\nCodigos,,Producto,Precio,Cantidades pedidas,,Talles/Colores\n6510,,Dino,9505,2,,';
    const bufComma = new TextEncoder().encode(csvComma).buffer;
    const resComma = parsearExcelLista006(bufComma, 'pedido_comma.csv');
    expect(resComma.ok).toBe(true);
    expect(resComma.lineas[0]?.cantidadTotal).toBe(2);
  });

  it('22. parsearMontoArgentino maneja formatos argentinos, typos y rechaza fechas y #VALUE!', () => {
    expect(parsearMontoArgentino('1.004.793,00')).toBe(1004793);
    expect(parsearMontoArgentino('690,573,06')).toBe(690573.06);
    expect(parsearMontoArgentino('$50.000')).toBe(50000);
    expect(parsearMontoArgentino('10/11')).toBe(0);
    expect(parsearMontoArgentino('#VALUE!')).toBe(0);
  });

  it('23. parsearFechaExcel corrige 2006 -> 2026 en seriales y texto', () => {
    expect(parsearFechaExcel('2006-11-24', '2026-01-01')).toBe('2026-11-24');
    expect(parsearFechaExcel('24/11/06', '2026-01-01')).toBe('2026-11-24');
  });

  it('24. importarPlanillaGeneral importa y deduplica por numeroPedido + cliente', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['N.º pedido', 'Cliente / negocio', 'WhatsApp', 'Fecha pedido', 'Total pedido', 'Monto abonado', 'Estado de pago', 'Armado', 'Envío'],
      ['0010', 'Pañalera Sur', '11223344', '20/11/2026', '150.000,00', '150.000,00', 'Pagado', 'Listo', 'Entregado'],
      ['0011', 'Juguetería Sol', '11556677', '21/11/2026', '200.000,00', '0', 'Cheque', 'Por armar', 'Pendiente']
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'PEDIDOS MAYORISTA MIBI');
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;

    const r1 = importarPlanillaGeneral(buf, 'planilla.xlsx', []);
    expect(r1.ok).toBe(true);
    expect(r1.nuevos).toHaveLength(2);
    expect(r1.nuevos[0]?.estadoPago).toBe('Pagado');
    expect(r1.nuevos[1]?.medioPago).toBe('Cheque');

    const r2 = importarPlanillaGeneral(buf, 'planilla.xlsx', r1.nuevos);
    expect(r2.nuevos).toHaveLength(0);
    expect(r2.duplicadosOmitidos).toBe(2);
  });

  // 25-28: Exportadores Excel (Hoja depósito, Picking, Tablero Maestro JSZip)
  it('25. Hoja de Depósito bloquea pedidos con talle/color sin definir o sin líneas', () => {
    const monster = CATALOGO_BY_CODIGO.get('6502')!;
    const pInc = recalcularPedido(makePedidoBase({ lineas: [resolverLinea(monster, 3, '')] }));
    const res = generarBufferHojaDeposito(pInc);
    expect(res.success).toBe(false);

    const pSinLineas = recalcularPedido(makePedidoBase({ lineas: [] }));
    expect(generarBufferHojaDeposito(pSinLineas).success).toBe(false);
  });

  it('26. Hoja de Depósito genera .xlsx válido con celdas numéricas reales', () => {
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const pOk = recalcularPedido(makePedidoBase({ lineas: [resolverLinea(dino, 4, '')], descuentoPorc: 10 }));
    const res = generarBufferHojaDeposito(pOk);
    expect(res.success).toBe(true);
    expect(res.buffer).toBeDefined();
    const wb = XLSX.read(res.buffer!, { type: 'array' });
    expect(wb.SheetNames[0]).toBe('Pedido_0001');
  });

  it('27. Picking consolidado excluye pedidos incompletos y consolida SKUs válidos', () => {
    const monster = CATALOGO_BY_CODIGO.get('6502')!;
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const pInc = recalcularPedido(makePedidoBase({ numeroPedido: '0001', lineas: [resolverLinea(monster, 2, '')] }));
    const pOk1 = recalcularPedido(makePedidoBase({ numeroPedido: '0002', lineas: [resolverLinea(dino, 3, '')] }));
    const pOk2 = recalcularPedido(makePedidoBase({ numeroPedido: '0003', lineas: [resolverLinea(dino, 5, '')] }));

    const soloInc = generarBufferPickingConsolidado([pInc]);
    expect(soloInc.success).toBe(false);
    expect(soloInc.excluidosTalle).toContain('#0001');

    const mix = generarBufferPickingConsolidado([pInc, pOk1, pOk2]);
    expect(mix.success).toBe(true);
    expect(mix.incluidos).toBe(2);
    expect(mix.totalUnidades).toBe(8);
    expect(mix.excluidosTalle).toEqual(['#0001']);
  });

  it('28. Tablero Maestro preserva ambas hojas oficiales de plantilla_pedidos_mibi.xlsx', async () => {
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const p = recalcularPedido(makePedidoBase({ lineas: [resolverLinea(dino, 2, '')] }));
    const { buffer } = await generarBufferTableroMaestro([p], calcularResumenFinanciero([p]));
    const zip = await JSZip.loadAsync(buffer);
    const s1 = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
    const s2 = await zip.file('xl/worksheets/sheet2.xml')?.async('string');
    expect(s1).toBeDefined();
    expect(s2).toBeDefined();
    expect(s1).toContain('<row r="11"');
    expect(s2).toContain('<row r="4"');
  });

  // 29-31: WhatsApp, Auth scrypt y Suite QA
  it('29. Mensajes WhatsApp de reclamo de talles y aviso de envío', () => {
    const monster = CATALOGO_BY_CODIGO.get('6502')!;
    const p = recalcularPedido(
      makePedidoBase({
        numeroPedido: '0042',
        cliente: 'Playa Kids',
        transporte: 'Andreani',
        nroSeguimiento: 'AND-9981',
        lineas: [resolverLinea(monster, 3, '')]
      })
    );
    const msgTalle = generarMensajeWhatsApp(p.cliente, p.numeroPedido, p.lineas);
    expect(msgTalle).toContain('Playa Kids');
    expect(msgTalle).toContain('#0042');
    expect(msgTalle).toContain('6502');

    const msgEnvio = generarMensajeDespachoWhatsApp(p);
    expect(msgEnvio).toContain('Andreani');
    expect(msgEnvio).toContain('AND-9981');
  });

  it('30. Hash de contraseñas scrypt + timingSafeEqual valida y rechaza correctamente', () => {
    const h = hashPassword('ClaveSegura2026!');
    expect(h.startsWith('scrypt$v1$')).toBe(true);
    expect(verifyPassword('ClaveSegura2026!', h)).toBe(true);
    expect(verifyPassword('otra-clave', h)).toBe(false);
  });

  it('31. ejecutarSuiteQA in-app aprueba el 100% de sus verificaciones', () => {
    const qa = ejecutarSuiteQA();
    expect(qa.length).toBeGreaterThanOrEqual(27);
    expect(qa.every((r) => r.ok)).toBe(true);
  });

  // 32-35: DAL + RBAC (VENTAS_DEPOSITO vs ADMIN) y concurrencia / soft-delete
  it('32. RBAC DAL: VENTAS_DEPOSITO NO puede crear pedidos pagados/con descuento ni modificar pagos', async () => {
    await reiniciarSistemaDal();
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000002',
      username: 'juancho',
      nombre: 'Juancho',
      rol: 'VENTAS_DEPOSITO',
      tokenVersion: 1
    };

    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    await expect(
      crearPedidoDal({
        cliente: 'Test',
        vendedor: 'Juancho',
        whatsapp: '',
        fechaPedido: getTodayISO(),
        despachoPrevisto: addDaysISO(getTodayISO(), 5),
        observaciones: '',
        descuentoPorc: 0,
        medioPago: 'Transferencia / Efectivo',
        montoAbonado: 1000,
        montoCheque: 0,
        fechaRecepcionCheque: '',
        fechaAcreditacionCheque: '',
        detalleCheque: '',
        lineas: [resolverLinea(dino, 2, '')]
      })
    ).rejects.toBeInstanceOf(AuthError);

    const creado = await crearPedidoDal({
      cliente: 'Test Ventas',
      vendedor: 'Juancho',
      whatsapp: '',
      fechaPedido: getTodayISO(),
      despachoPrevisto: addDaysISO(getTodayISO(), 5),
      observaciones: '',
      descuentoPorc: 0,
      medioPago: 'Sin pagar',
      montoAbonado: 0,
      montoCheque: 0,
      fechaRecepcionCheque: '',
      fechaAcreditacionCheque: '',
      detalleCheque: '',
      lineas: [resolverLinea(dino, 2, '')]
    });
    expect(creado.estadoPago).toBe('Falta pago');

    await expect(patchPedidoDal(creado.id, { montoAbonado: 5000 }, 'Hack')).rejects.toBeInstanceOf(AuthError);
    await expect(
      batchUpdatePedidosDal([creado.id], { type: 'pago', estado: 'Pagado' })
    ).rejects.toBeInstanceOf(AuthError);
    await expect(softDeletePedidosDal([creado.id])).rejects.toBeInstanceOf(AuthError);
  });

  it('33. DAL: actualización parcial atómica no pisa campos concurrentes (Depósito vs Admin)', async () => {
    await reiniciarSistemaDal();
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const creado = await crearPedidoDal({
      cliente: 'Concurrencia SA',
      vendedor: 'Juancho',
      whatsapp: '',
      fechaPedido: getTodayISO(),
      despachoPrevisto: addDaysISO(getTodayISO(), 5),
      observaciones: '',
      descuentoPorc: 0,
      medioPago: 'Sin pagar',
      montoAbonado: 0,
      montoCheque: 0,
      fechaRecepcionCheque: '',
      fechaAcreditacionCheque: '',
      detalleCheque: '',
      lineas: [resolverLinea(dino, 2, '')]
    });

    // Depósito actualiza transporte y guía
    const trasDeposito = await patchPedidoDal(
      creado.id,
      { transporte: 'Andreani', nroSeguimiento: 'AND-777', estadoEnvio: 'Despachado' },
      'Despachado'
    );
    // Admin acredita pago sin enviar transporte ni guía
    const trasAdmin = await patchPedidoDal(
      creado.id,
      { medioPago: 'Transferencia / Efectivo', montoAbonado: trasDeposito.totalFacturado, chequeAcreditado: true },
      'Cobrado'
    );

    expect(trasAdmin.transporte).toBe('Andreani');
    expect(trasAdmin.nroSeguimiento).toBe('AND-777');
    expect(trasAdmin.estadoEnvio).toBe('Despachado');
    expect(trasAdmin.estadoPago).toBe('Pagado');
    expect(trasAdmin.version).toBe(3);
  });

  it('34. DAL: Soft-delete y Restore soportan Deshacer de 6 segundos', async () => {
    await reiniciarSistemaDal();
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const creado = await crearPedidoDal({
      cliente: 'SoftDelete SRL',
      vendedor: 'Rocío',
      whatsapp: '',
      fechaPedido: getTodayISO(),
      despachoPrevisto: addDaysISO(getTodayISO(), 5),
      observaciones: '',
      descuentoPorc: 0,
      medioPago: 'Sin pagar',
      montoAbonado: 0,
      montoCheque: 0,
      fechaRecepcionCheque: '',
      fechaAcreditacionCheque: '',
      detalleCheque: '',
      lineas: [resolverLinea(dino, 1, '')]
    });

    expect(await softDeletePedidosDal([creado.id])).toBe(1);
    expect(await restorePedidosDal([creado.id])).toBe(1);
  });

  it('35. DAL: Server recalcula precios desde catálogo oficial ignorando manipulación de precio cliente', async () => {
    await reiniciarSistemaDal();
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const lineaManipulada = { ...resolverLinea(dino, 2, ''), precioUnitarioNeto: 1 };
    const creado = await crearPedidoDal({
      cliente: 'AntiTamper',
      vendedor: 'Juancho',
      whatsapp: '',
      fechaPedido: getTodayISO(),
      despachoPrevisto: addDaysISO(getTodayISO(), 5),
      observaciones: '',
      descuentoPorc: 0,
      medioPago: 'Sin pagar',
      montoAbonado: 0,
      montoCheque: 0,
      fechaRecepcionCheque: '',
      fechaAcreditacionCheque: '',
      detalleCheque: '',
      lineas: [lineaManipulada]
    });
    expect(creado.subtotalNeto).toBe(2 * 9505);
  });
});
