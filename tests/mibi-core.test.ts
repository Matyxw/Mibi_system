import fs from 'node:fs';
import path from 'node:path';
import type { NextRequest } from 'next/server';
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
  aplicarLineaAlBorrador,
  aplicarResolucionLinea,
  crearLineasEjemploLista006,
  importarPlanillaGeneral,
  limpiarTextoInvisible,
  parsearCantidadEstricta,
  parsearExcelLista006,
  parsearFechaExcel,
  parsearMontoArgentino,
  resolverLinea,
  sugerirProductosCercanos
} from '@/lib/excel-parser';
import {
  generarBufferHojaDeposito,
  generarBufferPickingConsolidado,
  generarBufferTableroMaestro
} from '@/lib/excel-exporter';
import {
  buildWhatsAppUrl,
  generarMensajeAjusteStockWhatsApp,
  generarMensajeDespachoWhatsApp,
  generarMensajeWhatsApp
} from '@/lib/whatsapp';
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
  actualizarArticulosPedidoDal,
  batchUpdatePedidosDal,
  ConflictError,
  crearPedidoDal,
  enviarPedidosACuentasDal,
  listarPedidosDal,
  marcarPedidoVistoDal,
  obtenerPedidoPorIdDal,
  obtenerPedidosPorIdsDal,
  patchPedidoDal,
  reiniciarSistemaDal,
  reportarIncidenciaDepositoDal,
  resolverIncidenciaDepositoDal,
  restorePedidosDal,
  retirarAsignacionDal,
  softDeletePedidosDal
} from '@/lib/dal/pedidos';
import { cambiarRolUsuarioDal, toggleActivoUsuarioDal } from '@/lib/dal/usuarios';
import {
  crearProductoCatalogoDal,
  customRecordToProductoCatalogo,
  editarProductoCatalogoDal,
  getCatalogoCompletoDal,
  toggleArchivoProductoCatalogoDal
} from '@/lib/dal/catalogo';
import { POST as exportTableroPost } from '@/app/api/export/tablero/route';
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

  it('36. DAL: permite editar cliente y whatsapp desde el panel de pedido sin alterar importes ni perder historial', async () => {
    await reiniciarSistemaDal();
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const creado = await crearPedidoDal({
      cliente: 'Cliente Original',
      vendedor: 'Juancho',
      whatsapp: '11-1111-1111',
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

    const editado = await patchPedidoDal(
      creado.id,
      { cliente: '  Juguetería Nueva SRL  ', whatsapp: '11-5606-9384' },
      'Actualizó cliente: Cliente Original → Juguetería Nueva SRL'
    );

    expect(editado.cliente).toBe('Juguetería Nueva SRL');
    expect(editado.whatsapp).toBe('11-5606-9384');
    expect(editado.totalFacturado).toBe(creado.totalFacturado);
    expect(editado.version).toBe(2);
    expect(editado.historial[0]?.accion).toContain('Juguetería Nueva SRL');
  });

  // 37-39: Evoluciones Modo Búnker (Hoja Depósito ciega, Catálogo extensible Pilusos, Derivación + Zero-Leak)
  it('37. Hoja de Depósito 100% ciega de precios: inspección celda por celda, 4 columnas y bloqueo por faltantes', () => {
    const monster = CATALOGO_BY_CODIGO.get('6502')!;
    const ejemploLineas = crearLineasEjemploLista006();
    const resolvedLineas = ejemploLineas.map((l) =>
      l.requiereRevision ? aplicarResolucionLinea(l, 2, 1) ?? l : l
    );

    const pedidoConPlata = recalcularPedido(
      makePedidoBase({
        numeroPedido: '0099',
        cliente: 'Mayorista Centro',
        vendedor: 'Agustín',
        transporte: 'Andreani',
        descuentoPorc: 15,
        medioPago: 'Transferencia / Efectivo',
        montoAbonado: 1500000,
        montoCheque: 0,
        lineas: resolvedLineas,
        asignaciones: [
          {
            id: 'asig_1',
            pedidoId: 'ped_test_1',
            usuarioId: '00000000-0000-4000-8000-000000000006',
            usuarioNombre: 'Depósito',
            usuarioRol: 'DEPOSITO',
            enviadoPor: '00000000-0000-4000-8000-000000000001',
            enviadoPorNombre: 'Administración',
            notaEnvio: 'Embalar en cajas dobles',
            prioridad: 'URGENTE',
            vistoAt: null,
            requiereRelectura: false,
            createdAt: Date.now(),
            updatedAt: Date.now()
          }
        ]
      })
    );

    expect(pedidoConPlata.totalFacturado).toBeGreaterThan(1000000);
    const exp = generarBufferHojaDeposito(pedidoConPlata);
    expect(exp.success).toBe(true);
    expect(exp.buffer).toBeDefined();

    const wb = XLSX.read(exp.buffer!, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]!]!;
    const ref = ws['!ref'];
    expect(ref).toBeDefined();
    const range = XLSX.utils.decode_range(ref!);

    const filas = XLSX.utils.sheet_to_json<(string | number | undefined)[]>(ws, { header: 1 });
    const headerIdx = filas.findIndex(
      (f) => f[0] === 'Código' && f[1] === 'Producto' && f[2] === 'Talle / Color' && f[3] === 'Cantidad pedida'
    );
    expect(headerIdx).toBeGreaterThanOrEqual(0);
    expect(filas[headerIdx]?.filter((x) => x !== undefined && x !== '')).toHaveLength(4);

    const totalUnidadesEsperado = pedidoConPlata.lineas.reduce((acc, l) => acc + l.cantidadTotal, 0);
    const filaTotal = filas.find((f) => f[2] === 'TOTAL UNIDADES');
    expect(filaTotal).toBeDefined();
    expect(filaTotal?.[3]).toBe(totalUnidadesEsperado);

    const terminosProhibidos = [
      '$',
      'ARS',
      'Precio',
      'Subtotal',
      'IVA',
      'Descuento',
      'Total B',
      'Facturado',
      'Cobrado',
      'Saldo'
    ];
    const montosProhibidos = new Set<number>([
      pedidoConPlata.subtotalNeto,
      pedidoConPlata.montoDescuento,
      pedidoConPlata.baseImponible,
      pedidoConPlata.iva21,
      pedidoConPlata.totalFacturado,
      pedidoConPlata.montoAbonado,
      pedidoConPlata.saldoPendiente,
      ...pedidoConPlata.lineas.map((l) => l.precioUnitarioNeto)
    ]);

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr] as XLSX.CellObject | undefined;
        if (!cell) continue;
        if (cell.z !== undefined && String(cell.z).includes('$')) {
          throw new Error(`Formato moneda prohibido en celda ${addr}: ${String(cell.z)}`);
        }
        const texto = String(cell.v ?? '');
        for (const term of terminosProhibidos) {
          if (texto.toLowerCase().includes(term.toLowerCase())) {
            throw new Error(`Texto financiero prohibido "${term}" en celda ${addr}: ${texto}`);
          }
        }
        if (typeof cell.v === 'number' && montosProhibidos.has(round2(cell.v))) {
          throw new Error(`Importe numérico prohibido filtrado en celda ${addr}: ${cell.v}`);
        }
      }
    }

    const pedidoIncompleto = recalcularPedido(
      makePedidoBase({ lineas: [resolverLinea(monster, 4, '')] })
    );
    expect(generarBufferHojaDeposito(pedidoIncompleto).success).toBe(false);
  });

  it('38. Catálogo extensible fuera de planilla (caso Pilusos): alta, cálculo, N variantes, Excel, WhatsApp y RBAC', async () => {
    const localDir = path.join(process.cwd(), '.local-data');
    if (fs.existsSync(localDir)) {
      fs.rmSync(localDir, { recursive: true, force: true });
    }

    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
    await reiniciarSistemaDal();

    // DEPOSITO no puede dar de alta productos en catálogo
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000006',
      username: 'deposito',
      nombre: 'Depósito',
      rol: 'DEPOSITO',
      tokenVersion: 1
    };
    await expect(
      crearProductoCatalogoDal({
        nombre: 'Piluso Bloqueado',
        rubro: 'PILUSOS',
        precioNeto: 7000,
        variantes: [{ codigo: 'PIL-BLK', etiqueta: 'Único' }]
      })
    ).rejects.toBeInstanceOf(AuthError);

    // VENTAS puede crear producto fuera de planilla con 1 o 3 variantes
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000002',
      username: 'juancho',
      nombre: 'Juancho',
      rol: 'VENTAS',
      tokenVersion: 1
    };

    const pilusoUnicoRec = await crearProductoCatalogoDal({
      nombre: 'Piluso Liso Beige',
      rubro: 'PILUSOS',
      precioNeto: 7500,
      variantes: [{ etiqueta: 'Único' }]
    });
    expect(pilusoUnicoRec.variantes[0]?.codigo).toMatch(/^EXT-\d{3}$/);

    const pilusoTallesRec = await crearProductoCatalogoDal({
      nombre: 'Piluso Shark',
      rubro: 'PILUSOS',
      precioNeto: 8500,
      variantes: [
        { codigo: 'PIL-S', etiqueta: 'T1 (1-3 años)' },
        { codigo: 'PIL-M', etiqueta: 'T2 (3-6 años)' },
        { codigo: 'PIL-L', etiqueta: 'T3 (6-9 años)' }
      ]
    });

    // VENTAS NO puede editar ni archivar producto existente del catálogo
    await expect(
      editarProductoCatalogoDal(pilusoTallesRec.id, {
        nombre: 'Hack',
        rubro: 'PILUSOS',
        precioNeto: 1,
        variantes: pilusoTallesRec.variantes
      })
    ).rejects.toBeInstanceOf(AuthError);
    await expect(toggleArchivoProductoCatalogoDal(pilusoTallesRec.id, false)).rejects.toBeInstanceOf(AuthError);

    const catalogoUnificado = await getCatalogoCompletoDal();
    expect(catalogoUnificado.length).toBeGreaterThanOrEqual(CATALOGO_MIBI.length + 2);

    const pilusoCat = customRecordToProductoCatalogo(pilusoTallesRec);
    const lPend = resolverLinea(pilusoCat, 6, 'surtido en 3 talles');
    expect(lPend.requiereRevision).toBe(true);

    const msgWa = generarMensajeWhatsApp('Tienda Pilusos', '0077', [lPend], catalogoUnificado);
    expect(msgWa).toContain('PIL-S');
    expect(msgWa).toContain('PIL-M');
    expect(msgWa).toContain('PIL-L');

    const lResuelta = aplicarResolucionLinea(lPend, [2, 3, 1], catalogoUnificado);
    expect(lResuelta).not.toBeNull();
    const lResueltaOk = lResuelta!;
    expect(lResueltaOk.requiereRevision).toBe(false);
    expect(lResueltaOk.desglose).toHaveLength(3);

    // Excel Lista 006 también reconoce códigos custom cuando recibe el catálogo unificado
    const wbCustom = XLSX.utils.book_new();
    const wsCustom = XLSX.utils.aoa_to_sheet([
      ['MIBI — LISTA N° 006'],
      ['CODIGOS', '', 'PRODUCTO', 'PRECIO NETO', 'CANTIDADES PEDIDAS', 'SUBTOTAL', 'TALLES / COLORES'],
      ['PIL-S', 'PIL-M', 'Piluso Shark', 8500, 6, '', '2 PIL-S, 3 PIL-M, 1 PIL-L']
    ]);
    XLSX.utils.book_append_sheet(wbCustom, wsCustom, 'Lista');
    const customBuf = XLSX.write(wbCustom, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
    const parsedExcel = parsearExcelLista006(customBuf, 'lista-custom.xlsx', catalogoUnificado);
    expect(parsedExcel.ok).toBe(true);
    expect(parsedExcel.lineas[0]?.cantidadTotal).toBe(6);
    expect(parsedExcel.lineas[0]?.requiereRevision).toBe(false);

    // Creación en DAL con cálculo exacto + Hoja Depósito + Picking consolidado
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
    const pedPiluso = await crearPedidoDal({
      cliente: 'Tienda Sur',
      vendedor: 'Juancho',
      whatsapp: '',
      fechaPedido: getTodayISO(),
      despachoPrevisto: addDaysISO(getTodayISO(), 4),
      observaciones: '',
      descuentoPorc: 10,
      medioPago: 'Sin pagar',
      montoAbonado: 0,
      montoCheque: 0,
      fechaRecepcionCheque: '',
      fechaAcreditacionCheque: '',
      detalleCheque: '',
      lineas: [lResueltaOk]
    });

    const baseEsperada = round2(round2(6 * 8500) * 0.9);
    const totalEsperado = round2(baseEsperada * 1.21);
    expect(pedPiluso.baseImponible).toBe(baseEsperada);
    expect(pedPiluso.totalFacturado).toBe(totalEsperado);

    const depBuf = generarBufferHojaDeposito(pedPiluso);
    expect(depBuf.success).toBe(true);
    const pickBuf = generarBufferPickingConsolidado([pedPiluso]);
    expect(pickBuf.success).toBe(true);
    expect(pickBuf.totalUnidades).toBe(6);

    // Editar precio futuro en catálogo no altera pedido histórico ya guardado
    await editarProductoCatalogoDal(pilusoTallesRec.id, {
      nombre: 'Piluso Shark',
      rubro: 'PILUSOS',
      precioNeto: 9900,
      variantes: pilusoTallesRec.variantes
    });
    const pedHistorico = await obtenerPedidoPorIdDal(pedPiluso.id);
    expect(pedHistorico?.totalFacturado).toBe(totalEsperado);
  });

  it('39. Derivación por cuenta (ADMIN -> DEPOSITO / VENTAS), compuertas, relectura y Zero-Leak servidor', async () => {
    const localDir = path.join(process.cwd(), '.local-data');
    if (fs.existsSync(localDir)) {
      fs.rmSync(localDir, { recursive: true, force: true });
    }

    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
    await reiniciarSistemaDal();

    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const monster = CATALOGO_BY_CODIGO.get('6502')!;

    const pCompletoSinPago = await crearPedidoDal({
      cliente: 'Cliente Norte',
      vendedor: 'Juancho',
      whatsapp: '1155550000',
      fechaPedido: getTodayISO(),
      despachoPrevisto: addDaysISO(getTodayISO(), 5),
      observaciones: 'Seña pendiente · transferencia CBU',
      descuentoPorc: 10,
      medioPago: 'Sin pagar',
      montoAbonado: 0,
      montoCheque: 0,
      fechaRecepcionCheque: '',
      fechaAcreditacionCheque: '',
      detalleCheque: '',
      lineas: [resolverLinea(dino, 4, '')]
    });

    const pIncompleto = await crearPedidoDal({
      cliente: 'Cliente Incompleto',
      vendedor: 'Pablo',
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
      lineas: [resolverLinea(monster, 3, '')]
    });

    const DEP_ID = '00000000-0000-4000-8000-000000000006';
    const JUANCHO_ID = '00000000-0000-4000-8000-000000000002';

    // Antes de derivar, DEPOSITO no ve ningún pedido
    mockSession = {
      userId: DEP_ID,
      username: 'deposito',
      nombre: 'Depósito',
      rol: 'DEPOSITO',
      tokenVersion: 1
    };
    expect(await listarPedidosDal()).toHaveLength(0);
    await expect(obtenerPedidoPorIdDal(pCompletoSinPago.id)).rejects.toBeInstanceOf(AuthError);

    // Compuerta Tablero Maestro bloqueada para DEPOSITO y VENTAS
    const resTableroDep = await exportTableroPost(
      new Request('http://localhost/api/export/tablero', { method: 'POST', body: '{}' }) as unknown as NextRequest
    );
    expect(resTableroDep.status).toBe(403);

    // ADMIN intenta enviar pedido incompleto a Depósito -> bloqueo duro
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
    const resBloqueo = await enviarPedidosACuentasDal({
      pedidoIds: [pIncompleto.id],
      usuarioIds: [DEP_ID],
      prioridad: 'NORMAL',
      notaEnvio: ''
    });
    expect(resBloqueo.ok).toBe(false);
    expect(resBloqueo.omitidosTalle).toContain(`#${pIncompleto.numeroPedido}`);

    // Enviar incompleto a VENTAS sí está permitido
    const resVentas = await enviarPedidosACuentasDal({
      pedidoIds: [pIncompleto.id],
      usuarioIds: [JUANCHO_ID],
      prioridad: 'NORMAL',
      notaEnvio: 'Pedir talle al cliente'
    });
    expect(resVentas.ok).toBe(true);

    // Enviar pedido completo sin pago a DEPOSITO exige confirmación explícita
    const resWarnPago = await enviarPedidosACuentasDal({
      pedidoIds: [pCompletoSinPago.id],
      usuarioIds: [DEP_ID],
      prioridad: 'URGENTE',
      notaEnvio: 'Prioridad camión 15hs'
    });
    expect(resWarnPago.ok).toBe(false);
    expect(resWarnPago.requiereConfirmacionSinPago).toBe(true);

    // Confirmando sin pago sí envía a DEPOSITO
    const resOkDep = await enviarPedidosACuentasDal({
      pedidoIds: [pCompletoSinPago.id],
      usuarioIds: [DEP_ID],
      prioridad: 'URGENTE',
      notaEnvio: 'Prioridad camión 15hs',
      forzarSinPago: true
    });
    expect(resOkDep.ok).toBe(true);
    expect(resOkDep.enviados).toBe(1);

    // Sesión DEPOSITO: solo ve el pedido asignado y con CERO datos monetarios (Zero-Leak)
    mockSession = {
      userId: DEP_ID,
      username: 'deposito',
      nombre: 'Depósito',
      rol: 'DEPOSITO',
      tokenVersion: 1
    };
    const listaDep = await listarPedidosDal();
    expect(listaDep).toHaveLength(1);
    const pedDep = listaDep[0]!;
    expect(pedDep.id).toBe(pCompletoSinPago.id);
    expect(pedDep.totalFacturado).toBe(0);
    expect(pedDep.subtotalNeto).toBe(0);
    expect(pedDep.iva21).toBe(0);
    expect(pedDep.descuentoPorc).toBe(0);
    expect(pedDep.saldoPendiente).toBe(0);
    expect(pedDep.lineas[0]?.precioUnitarioNeto).toBe(0);

    // Marcar visto + reportar incidencia desde Depósito
    await marcarPedidoVistoDal(pCompletoSinPago.id);
    const pedPostVisto = await obtenerPedidoPorIdDal(pCompletoSinPago.id);
    expect(pedPostVisto?.asignaciones.find((a) => a.usuarioId === DEP_ID)?.vistoAt).not.toBeNull();

    const pedIncidencia = await reportarIncidenciaDepositoDal(pCompletoSinPago.id, 'Falta 1 caja de Dino');
    expect(pedIncidencia.incidenciaDeposito).toBe('Falta 1 caja de Dino');

    // ADMIN edita cantidades del pedido ya enviado -> dispara requiereRelectura y resetea vistoAt
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
    await resolverIncidenciaDepositoDal(pCompletoSinPago.id);
    await patchPedidoDal(
      pCompletoSinPago.id,
      { lineas: [resolverLinea(dino, 6, '')] },
      'Ajustó cantidad de 4 a 6 unidades'
    );

    mockSession = {
      userId: DEP_ID,
      username: 'deposito',
      nombre: 'Depósito',
      rol: 'DEPOSITO',
      tokenVersion: 1
    };
    const pedRelectura = await obtenerPedidoPorIdDal(pCompletoSinPago.id);
    const asigDep = pedRelectura?.asignaciones.find((a) => a.usuarioId === DEP_ID);
    expect(asigDep?.requiereRelectura).toBe(true);
    expect(asigDep?.vistoAt).toBeNull();

    // ADMIN retira asignación -> desaparece de la bandeja de Depósito y bloquea acceso directo (403)
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
    await retirarAsignacionDal(pCompletoSinPago.id, DEP_ID);

    mockSession = {
      userId: DEP_ID,
      username: 'deposito',
      nombre: 'Depósito',
      rol: 'DEPOSITO',
      tokenVersion: 1
    };
    expect(await listarPedidosDal()).toHaveLength(0);
    await expect(obtenerPedidoPorIdDal(pCompletoSinPago.id)).rejects.toBeInstanceOf(AuthError);
    expect(await obtenerPedidosPorIdsDal([pCompletoSinPago.id])).toHaveLength(0);

    // VENTAS ve su pedido asignado y no accede al Tablero Maestro
    mockSession = {
      userId: JUANCHO_ID,
      username: 'juancho',
      nombre: 'Juancho',
      rol: 'VENTAS',
      tokenVersion: 1
    };
    const listaVentas = await listarPedidosDal();
    expect(listaVentas.some((p) => p.id === pIncompleto.id)).toBe(true);
    const resTableroVentas = await exportTableroPost(
      new Request('http://localhost/api/export/tablero', { method: 'POST', body: '{}' }) as unknown as NextRequest
    );
    expect(resTableroVentas.status).toBe(403);
  });

  // 40-42: Cirugía 360 (WhatsApp URL, máquina de estados coherente y anti-autobloqueo ADMIN)
  it('40. buildWhatsAppUrl normaliza teléfonos argentinos y descarta inválidos', () => {
    const uLocal = buildWhatsAppUrl('11-5606-9384', 'Hola Mibi');
    expect(uLocal).toBe('https://wa.me/5491156069384?text=Hola%20Mibi');

    const uCero = buildWhatsAppUrl('011 5606 9384', 'Consulta');
    expect(uCero).toBe('https://wa.me/5491156069384?text=Consulta');

    const uMas549 = buildWhatsAppUrl('+54 9 11 5606-9384', 'Guía');
    expect(uMas549).toBe('https://wa.me/5491156069384?text=Gu%C3%ADa');

    expect(buildWhatsAppUrl('sin tel', 'Hola')).toBeNull();
    expect(buildWhatsAppUrl('11-5606-9384', '   ')).toBeNull();
  });

  it('41. Máquina de estados coherente en servidor (armado/envío/fechas, cheques 100%, bloqueo y conflicto)', async () => {
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
    await reiniciarSistemaDal();

    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const monster = CATALOGO_BY_CODIGO.get('6502')!;

    // Bloqueo duro si faltantesVariante > 0
    const pedPendTalle = await crearPedidoDal({
      cliente: 'Bloqueo Talle',
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
      lineas: [resolverLinea(monster, 3, '')]
    });

    await expect(
      patchPedidoDal(pedPendTalle.id, { estadoEnvio: 'Despachado' }, 'Intento ilegal')
    ).rejects.toThrow();
    await expect(
      patchPedidoDal(pedPendTalle.id, { estadoEnvio: 'Entregado' }, 'Intento ilegal')
    ).rejects.toThrow();
    await expect(
      patchPedidoDal(pedPendTalle.id, { estadoArmado: 'Listo' }, 'Intento ilegal')
    ).rejects.toThrow();

    // Pedido completo: Despachado promueve Por armar -> Listo y estampa fechaDespacho; retroceso limpia fechas
    const hoy = getTodayISO();
    const pedCompleto = await crearPedidoDal({
      cliente: 'Estado Coherente',
      vendedor: 'Juancho',
      whatsapp: '11-5606-9384',
      fechaPedido: hoy,
      despachoPrevisto: addDaysISO(hoy, 5),
      observaciones: '',
      descuentoPorc: 0,
      medioPago: 'Cheque',
      montoAbonado: 0,
      montoCheque: 2 * 9505 * 1.21,
      fechaRecepcionCheque: hoy,
      fechaAcreditacionCheque: addDaysISO(hoy, 30),
      detalleCheque: 'Banco Nación',
      lineas: [resolverLinea(dino, 2, '')]
    });

    expect(pedCompleto.estadoArmado).toBe('Por armar');

    const desp = await patchPedidoDal(
      pedCompleto.id,
      { estadoEnvio: 'Despachado' },
      'Despacho directo',
      pedCompleto.version
    );
    expect(desp.estadoArmado).toBe('Listo');
    expect(desp.fechaDespacho).toBe(hoy);
    expect(desp.fechaEntrega ?? '').toBe('');

    const entr = await patchPedidoDal(pedCompleto.id, { estadoEnvio: 'Entregado' }, 'Entrega', desp.version);
    expect(entr.fechaDespacho).toBe(hoy);
    expect(entr.fechaEntrega).toBe(hoy);

    const vueltoDesp = await patchPedidoDal(
      pedCompleto.id,
      { estadoEnvio: 'Despachado' },
      'Corrección entrega',
      entr.version
    );
    expect(vueltoDesp.fechaEntrega).toBe('');
    expect(vueltoDesp.fechaDespacho).toBe(hoy);

    const vueltoPend = await patchPedidoDal(
      pedCompleto.id,
      { estadoEnvio: 'Pendiente' },
      'Corrección envío',
      vueltoDesp.version
    );
    expect(vueltoPend.fechaDespacho).toBe('');
    expect(vueltoPend.fechaEntrega).toBe('');

    // Cambiar fechaRecepcionCheque recalcula +30 días automáticamente
    const fechaRec = '2026-06-01';
    const cheqRec = await patchPedidoDal(
      pedCompleto.id,
      { fechaRecepcionCheque: fechaRec },
      'Cambio fecha recepción'
    );
    expect(cheqRec.fechaAcreditacionCheque).toBe('2026-07-01');

    // Abonar 100% en pedido con Cheque pasa automáticamente a Pagado y chequeAcreditado = true
    const pagadoCheque = await patchPedidoDal(
      pedCompleto.id,
      { montoAbonado: cheqRec.totalFacturado },
      'Abonó 100%'
    );
    expect(pagadoCheque.estadoPago).toBe('Pagado');
    expect(pagadoCheque.chequeAcreditado).toBe(true);
    expect(pagadoCheque.montoCheque).toBe(0);

    // Detección de conflicto de concurrencia cuando expectedVersion no coincide
    await expect(
      patchPedidoDal(pedCompleto.id, { transporte: 'Andreani' }, 'Versión vieja', vueltoPend.version)
    ).rejects.toBeInstanceOf(ConflictError);
  });

  it('42. Seguridad de equipo: ADMIN no puede desactivarse ni quitarse su propio rol ADMIN', async () => {
    const ADMIN_ID = '00000000-0000-4000-8000-000000000001';
    mockSession = {
      userId: ADMIN_ID,
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };

    await expect(toggleActivoUsuarioDal(ADMIN_ID, false)).rejects.toThrow();
    await expect(cambiarRolUsuarioDal(ADMIN_ID, 'VENTAS')).rejects.toThrow();
  });

  // 43-48: Resolución guiada de alertas Excel & anti-typo & inmunidad contable
  it('43. Alerta enriquecida CODIGO_DESCONOCIDO extrae metadatos y nunca auto-crea ni factura el producto', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['MIBI LISTA N° 006'],
      ['Codigos', '', 'Producto', 'Precio', 'Cantidades pedidas', '', 'Talles/Colores'],
      ['9999', '', 'Piluso Tiburón Azul', 15000, 6, '', '3 T1, 3 T2']
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Lista');
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
    const res = parsearExcelLista006(buf, 'desconocido.xlsx');

    expect(res.ok).toBe(true);
    expect(res.lineas).toHaveLength(0);
    expect(res.advertencias).toHaveLength(1);
    const adv = res.advertencias[0]!;
    expect(adv.tipo).toBe('CODIGO_DESCONOCIDO');
    expect(adv.codigosDetectados).toEqual(['9999']);
    expect(adv.descripcionExcel).toBe('Piluso Tiburón Azul');
    expect(adv.precioExcelReferencia).toBe(15000);
    expect(adv.cantidadDetectada).toBe(6);
    expect(adv.notaDetectada).toBe('3 T1, 3 T2');
    expect(adv.productoExistenteIdGrupo).toBeNull();
  });

  it('44. Par desconocido A+B detecta ambos códigos en codigosDetectados', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['MIBI LISTA N° 006'],
      ['Codigos', '', 'Producto', 'Precio', 'Cantidades pedidas', '', 'Talles/Colores'],
      ['9801', '9802', 'Piluso Doble', 18000, 4, '', '2 y 2']
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
    const res = parsearExcelLista006(buf, 'par_desconocido.xlsx');

    expect(res.ok).toBe(true);
    expect(res.advertencias).toHaveLength(1);
    expect(res.advertencias[0]?.codigosDetectados).toEqual(['9801', '9802']);
    expect(res.advertencias[0]?.codigo).toBe('9801/9802');
    expect(res.advertencias[0]?.cantidadDetectada).toBe(4);
  });

  it('45. sugerirProductosCercanos detecta typos de código, nombres parecidos y devuelve [] sin coincidencia', () => {
    const porCodigo = sugerirProductosCercanos('650X', '', CATALOGO_MIBI);
    expect(porCodigo.length).toBeGreaterThanOrEqual(1);
    expect(porCodigo.length).toBeLessThanOrEqual(3);
    expect(porCodigo[0]?.motivo).toBe('Código parecido');

    const porNombre = sugerirProductosCercanos('8888', 'Chaleco Shark', CATALOGO_MIBI);
    expect(porNombre.length).toBeGreaterThanOrEqual(1);
    expect(porNombre[0]?.idGrupo).toBe('6504-6505');
    expect(porNombre[0]?.motivo).toBe('Nombre parecido');

    const sinMatch = sugerirProductosCercanos('9999', 'Juguete Orbital X', CATALOGO_MIBI);
    expect(sinMatch).toEqual([]);
  });

  it('46. CANTIDAD_INVALIDA registra valor crudo y productoExistenteIdGrupo si el código existe', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['MIBI LISTA N° 006'],
      ['Codigos', '', 'Producto', 'Precio', 'Cantidades pedidas', '', 'Talles/Colores'],
      ['6510', '', 'Antiparras Dino', 9505, 'abc', '', ''],
      ['9998', '', 'Producto Desconocido', 1200, '2 doc', '', '']
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
    const res = parsearExcelLista006(buf, 'cant_invalida.xlsx');

    expect(res.ok).toBe(true);
    expect(res.lineas).toHaveLength(0);
    expect(res.advertencias).toHaveLength(2);

    const aExist = res.advertencias[0]!;
    expect(aExist.tipo).toBe('CANTIDAD_INVALIDA');
    expect(aExist.cantidadDetectada).toBeNull();
    expect(aExist.cantidadRaw).toBe('abc');
    expect(aExist.productoExistenteIdGrupo).toBe('6510');

    const aDesc = res.advertencias[1]!;
    expect(aDesc.tipo).toBe('CANTIDAD_INVALIDA');
    expect(aDesc.cantidadRaw).toBe('2 doc');
    expect(aDesc.productoExistenteIdGrupo).toBeNull();
  });

  it('47. Inmunidad de precios y avisosPrecio: nunca factura con precio editado del Excel', () => {
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['MIBI LISTA N° 006'],
      ['Codigos', '', 'Producto', 'Precio', 'Cantidades pedidas', '', 'Talles/Colores'],
      ['6510', '', 'Antiparras Dino', 100, 2, '', ''],
      ['6511', '', 'Antiparras Monster', CATALOGO_BY_CODIGO.get('6511')!.precioNeto, 1, '', '']
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Lista');
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
    const res = parsearExcelLista006(buf, 'precio_editado.xlsx');

    expect(res.ok).toBe(true);
    expect(res.advertencias).toHaveLength(0);
    expect(res.lineas).toHaveLength(2);
    const lineaDino = res.lineas.find((l) => l.idGrupo === '6510')!;
    expect(lineaDino.precioUnitarioNeto).toBe(dino.precioNeto);
    expect(res.avisosPrecio).toHaveLength(1);
    expect(res.avisosPrecio?.[0]).toEqual({
      codigo: '6510',
      nombreProducto: dino.nombre,
      precioPlanilla: 100,
      precioVigente: dino.precioNeto
    });
  });

  it('48. aplicarLineaAlBorrador suma cantidades y protege desglose en productos multi-variante', () => {
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const shark = CATALOGO_BY_CODIGO.get('6504')!;

    // Caso 1: variante única fusiona cantidad
    const baseUnico = aplicarLineaAlBorrador([], dino, 2, '');
    const fusionUnico = aplicarLineaAlBorrador(baseUnico, dino, 3, '');
    expect(fusionUnico).toHaveLength(1);
    expect(fusionUnico[0]?.cantidadTotal).toBe(5);
    expect(fusionUnico[0]?.estado).toBe('OK');

    // Caso 2: multi-variante ya OK al sumar cantidad sin nota vuelve seguro a REQUIERE_REVISION
    const baseShark = aplicarLineaAlBorrador([], shark, 2, 'de 1 a 3');
    expect(baseShark[0]?.estado).toBe('OK');
    const sharkSinNota = aplicarLineaAlBorrador(baseShark, shark, 2, '');
    expect(sharkSinNota[0]?.cantidadTotal).toBe(4);
    expect(sharkSinNota[0]?.estado).toBe('REQUIERE_REVISION');
    expect(sharkSinNota[0]?.requiereRevision).toBe(true);

    // Caso 3: multi-variante ya OK al sumar cantidad con desglose explícito mantiene OK sumando por talle
    const sharkConNota = aplicarLineaAlBorrador(baseShark, shark, 3, 'de 3 a 6');
    expect(sharkConNota[0]?.cantidadTotal).toBe(5);
    expect(sharkConNota[0]?.estado).toBe('OK');
    expect(sharkConNota[0]?.desglose).toEqual([
      { codigo: '6504', etiqueta: 'T1 (1-3 años)', cantidad: 2 },
      { codigo: '6505', etiqueta: 'T2 (3-6 años)', cantidad: 3 }
    ]);
  });

  // 49-55 (Tests A-G): Edición de artículos guardados, precio histórico, saldo a favor, cheque, depósito, RBAC y WhatsApp
  it('49. Test A: Reducción de cantidad en variante específica conserva precio histórico y genera diff legible', async () => {
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };

    const shark = CATALOGO_BY_CODIGO.get('6504')!;
    const pilusoHist = await crearProductoCatalogoDal({
      nombre: 'Piluso Histórico Test A',
      rubro: 'PILUSOS',
      precioNeto: 8500,
      variantes: [{ codigo: 'PIL-HIST-1', etiqueta: 'Único', aliases: ['unico'] }]
    });

    const ped = await crearPedidoDal({
      cliente: 'Cliente Precio Histórico',
      vendedor: 'Juancho',
      whatsapp: '11-5555-1111',
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
      lineas: [
        resolverLinea(shark, 6, '3 t1, 3 t2'),
        resolverLinea(customRecordToProductoCatalogo(pilusoHist), 2, '')
      ]
    });

    // Subimos el precio en catálogo a $12.000 después de guardar el pedido
    await editarProductoCatalogoDal(pilusoHist.id, {
      nombre: 'Piluso Histórico Test A',
      rubro: 'PILUSOS',
      precioNeto: 12000,
      variantes: [{ codigo: 'PIL-HIST-1', etiqueta: 'Único', aliases: ['unico'] }]
    });

    // Bajamos Shark T1 de 3 a 1 (total Shark = 4 u.) sin tocar Piluso
    const lineaSharkEditada = {
      ...ped.lineas[0]!,
      cantidadTotal: 4,
      desglose: [
        { codigo: '6504', etiqueta: 'T1 (1-3 años)', cantidad: 1 },
        { codigo: '6505', etiqueta: 'T2 (3-6 años)', cantidad: 3 }
      ]
    };

    const res = await actualizarArticulosPedidoDal({
      pedidoId: ped.id,
      expectedVersion: ped.version,
      lineas: [lineaSharkEditada, ped.lineas[1]!]
    });

    const pilusoLineaPost = res.pedido.lineas.find((l) => l.idGrupo === pilusoHist.idGrupo)!;
    const sharkLineaPost = res.pedido.lineas.find((l) => l.idGrupo === shark.idGrupo)!;
    expect(sharkLineaPost.estado).toBe('OK');
    expect(sharkLineaPost.cantidadTotal).toBe(4);
    expect(pilusoLineaPost.precioUnitarioNeto).toBe(8500);

    // Recálculo exacto en centavos: (4 * 46612 + 2 * 8500) * 1.21
    const esperadoSubtotal = 4 * shark.precioNeto + 2 * 8500;
    const esperadoIva = round2(esperadoSubtotal * 0.21);
    expect(res.pedido.subtotalNeto).toBe(esperadoSubtotal);
    expect(res.pedido.iva21).toBe(esperadoIva);
    expect(res.pedido.totalFacturado).toBe(round2(esperadoSubtotal + esperadoIva));
    expect(res.pedido.historial.some((h) => h.accion.includes('3 → 1 u.'))).toBe(true);
  });

  it('50. Test B: Quitar línea sin stock y agregar reemplazo de productos_custom recalcula exacto', async () => {
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };

    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const pilusoNuevo = await crearProductoCatalogoDal({
      nombre: 'Piluso Reemplazo B',
      rubro: 'PILUSOS',
      precioNeto: 9000,
      variantes: [{ codigo: 'PIL-REEM-B', etiqueta: 'Único', aliases: ['unico'] }]
    });

    const ped = await crearPedidoDal({
      cliente: 'Cliente Reemplazo',
      vendedor: 'Juancho',
      whatsapp: '11-5555-2222',
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
      lineas: [resolverLinea(dino, 3, '')]
    });

    const res = await actualizarArticulosPedidoDal({
      pedidoId: ped.id,
      expectedVersion: ped.version,
      lineas: [resolverLinea(customRecordToProductoCatalogo(pilusoNuevo), 2, '')]
    });

    expect(res.pedido.lineas).toHaveLength(1);
    expect(res.pedido.totalFacturado).toBe(round2(2 * 9000 * 1.21));
    expect(res.resumenAjuste.cambios.some((c) => c.tipo === 'quitado' && c.nombreProducto === dino.nombre)).toBe(true);
    expect(res.resumenAjuste.cambios.some((c) => c.tipo === 'agregado' && c.nombreProducto === 'Piluso Reemplazo B')).toBe(
      true
    );
    expect(res.pedido.historial.some((h) => h.accion.includes('Quitó:') && h.accion.includes('Agregó:'))).toBe(true);
  });

  it('51. Test C: Caso borde Saldo a Favor cuando montoAbonado supera el nuevo total', async () => {
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };

    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const qtyInicial = 10;
    const totalInicial = round2(qtyInicial * dino.precioNeto * 1.21);

    const ped = await crearPedidoDal({
      cliente: 'Cliente Saldo Favor',
      vendedor: 'Juancho',
      whatsapp: '11-5555-3333',
      fechaPedido: getTodayISO(),
      despachoPrevisto: addDaysISO(getTodayISO(), 5),
      observaciones: 'Pago adelantado completo',
      descuentoPorc: 0,
      medioPago: 'Transferencia / Efectivo',
      montoAbonado: totalInicial,
      montoCheque: 0,
      fechaRecepcionCheque: '',
      fechaAcreditacionCheque: '',
      detalleCheque: '',
      lineas: [resolverLinea(dino, qtyInicial, '')]
    });

    const res = await actualizarArticulosPedidoDal({
      pedidoId: ped.id,
      expectedVersion: ped.version,
      lineas: [resolverLinea(dino, 6, '')]
    });

    const nuevoTotalEsperado = round2(6 * dino.precioNeto * 1.21);
    const saldoFavorEsperado = round2(totalInicial - nuevoTotalEsperado);

    expect(res.pedido.totalFacturado).toBe(nuevoTotalEsperado);
    expect(res.pedido.montoAbonado).toBe(nuevoTotalEsperado);
    expect(res.pedido.estadoPago).toBe('Pagado');
    expect(res.pedido.saldoPendiente).toBe(0);
    expect(res.resumenAjuste.saldoAFavorGenerado).toBe(saldoFavorEsperado);
    expect(res.pedido.observaciones).toContain('Saldo a favor del cliente por ajuste de stock');
    expect(res.pedido.historial.some((h) => h.accion.includes('Saldo a favor del cliente'))).toBe(true);
  });

  it('52. Test D: Pedido Pagado que sube pasa a Pendiente y Cheque pendiente acota montoCheque si baja', async () => {
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };

    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const tot2 = round2(2 * dino.precioNeto * 1.21);

    // Subcaso 1: Pedido Pagado aumenta de 2 a 4 u. -> vuelve a Pendiente con saldoPendiente exacto
    const pedPagado = await crearPedidoDal({
      cliente: 'Cliente Sube Total',
      vendedor: 'Juancho',
      whatsapp: '11-5555-4444',
      fechaPedido: getTodayISO(),
      despachoPrevisto: addDaysISO(getTodayISO(), 5),
      observaciones: '',
      descuentoPorc: 0,
      medioPago: 'Transferencia / Efectivo',
      montoAbonado: tot2,
      montoCheque: 0,
      fechaRecepcionCheque: '',
      fechaAcreditacionCheque: '',
      detalleCheque: '',
      lineas: [resolverLinea(dino, 2, '')]
    });

    const resSube = await actualizarArticulosPedidoDal({
      pedidoId: pedPagado.id,
      expectedVersion: pedPagado.version,
      lineas: [resolverLinea(dino, 4, '')]
    });
    const tot4 = round2(4 * dino.precioNeto * 1.21);
    expect(resSube.pedido.estadoPago).toBe('Pendiente');
    expect(resSube.pedido.saldoPendiente).toBe(round2(tot4 - tot2));

    // Subcaso 2: Pedido con Cheque pendiente baja de 4 a 2 u. -> acota montoCheque y reporta diferenciaCheque
    const pedCheque = await crearPedidoDal({
      cliente: 'Cliente Cheque Baja',
      vendedor: 'Juancho',
      whatsapp: '11-5555-5555',
      fechaPedido: getTodayISO(),
      despachoPrevisto: addDaysISO(getTodayISO(), 5),
      observaciones: '',
      descuentoPorc: 0,
      medioPago: 'Cheque',
      montoAbonado: 0,
      montoCheque: tot4,
      fechaRecepcionCheque: getTodayISO(),
      fechaAcreditacionCheque: addDaysISO(getTodayISO(), 30),
      detalleCheque: 'Galicia #901',
      lineas: [resolverLinea(dino, 4, '')]
    });

    const resCheque = await actualizarArticulosPedidoDal({
      pedidoId: pedCheque.id,
      expectedVersion: pedCheque.version,
      lineas: [resolverLinea(dino, 2, '')]
    });
    expect(resCheque.pedido.montoCheque).toBe(tot2);
    expect(resCheque.resumenAjuste.diferenciaCheque).toBe(round2(tot4 - tot2));
  });

  it('53. Test E: Sincronización con Depósito (requiereRelectura), cierre de incidencia y Hoja Depósito sin precios', async () => {
    const DEP_ID = '00000000-0000-4000-8000-000000000006';
    const dino = CATALOGO_BY_CODIGO.get('6510')!;
    const shark = CATALOGO_BY_CODIGO.get('6504')!;

    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };

    const ped = await crearPedidoDal({
      cliente: 'Cliente Depósito Sync',
      vendedor: 'Juancho',
      whatsapp: '11-5555-6666',
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
      lineas: [resolverLinea(shark, 4, '2 t1, 2 t2'), resolverLinea(dino, 3, '')]
    });

    await enviarPedidosACuentasDal({
      pedidoIds: [ped.id],
      usuarioIds: [DEP_ID],
      prioridad: 'NORMAL',
      notaEnvio: '',
      forzarSinPago: true
    });

    // Depósito marca como visto y luego reporta faltante
    mockSession = {
      userId: DEP_ID,
      username: 'deposito',
      nombre: 'Depósito',
      rol: 'DEPOSITO',
      tokenVersion: 1
    };
    await marcarPedidoVistoDal(ped.id);
    const conIncidencia = await reportarIncidenciaDepositoDal(ped.id, 'Faltan 2 Dino');

    // ADMIN ajusta cantidades y marca resolverIncidencia: true
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
    const res = await actualizarArticulosPedidoDal({
      pedidoId: ped.id,
      expectedVersion: conIncidencia.version,
      lineas: [ped.lineas[0]!, resolverLinea(dino, 1, '')],
      resolverIncidencia: true
    });

    expect(res.pedido.incidenciaDeposito).toBe('');
    expect(res.resumenAjuste.avisoRelecturaDeposito).toBe(true);
    expect(res.resumenAjuste.incidenciaResuelta).toBe(true);
    expect(res.pedido.asignaciones.some((a) => a.requiereRelectura && a.vistoAt === null)).toBe(true);

    // Verificar Hoja de Depósito sin precios y con nueva cantidad (4 + 1 = 5 u.)
    const hoja = generarBufferHojaDeposito(res.pedido);
    expect(hoja.success).toBe(true);
    const wb = XLSX.read(hoja.buffer!, { type: 'buffer' });
    const rows = XLSX.utils.sheet_to_json<Array<string | number>>(wb.Sheets[wb.SheetNames[0]!]!, { header: 1 });
    const textoHoja = rows.flat().join(' ');
    expect(textoHoja).not.toContain('$');
    const filaDino = rows.find((r) => String(r[0]) === '6510');
    expect(filaDino?.[3]).toBe(1);
    const filaTotal = rows.find((r) => String(r[2]).includes('TOTAL UNIDADES'));
    expect(filaTotal?.[3]).toBe(5);
  });

  it('54. Test F: Seguridad RBAC (DEPOSITO / VENTAS), anti-vacío y ConflictError', async () => {
    const DEP_ID = '00000000-0000-4000-8000-000000000006';
    const JUANCHO_ID = '00000000-0000-4000-8000-000000000002';
    const dino = CATALOGO_BY_CODIGO.get('6510')!;

    mockSession = {
      userId: JUANCHO_ID,
      username: 'juancho',
      nombre: 'Juancho',
      rol: 'VENTAS',
      tokenVersion: 1
    };
    const pedVentas = await crearPedidoDal({
      cliente: 'Cliente RBAC',
      vendedor: 'Juancho',
      whatsapp: '11-5555-7777',
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

    // VENTAS puede editar su pedido mientras está sin pagos y sin enviar a Depósito
    const okVentas = await actualizarArticulosPedidoDal({
      pedidoId: pedVentas.id,
      expectedVersion: pedVentas.version,
      lineas: [resolverLinea(dino, 3, '')]
    });
    expect(okVentas.pedido.version).toBe(pedVentas.version + 1);

    // Concurrencia: expectedVersion viejo lanza ConflictError
    await expect(
      actualizarArticulosPedidoDal({
        pedidoId: pedVentas.id,
        expectedVersion: pedVentas.version,
        lineas: [resolverLinea(dino, 5, '')]
      })
    ).rejects.toBeInstanceOf(ConflictError);

    // Anti-vacío: 0 artículos es rechazado en servidor
    await expect(
      actualizarArticulosPedidoDal({
        pedidoId: pedVentas.id,
        expectedVersion: okVentas.pedido.version,
        lineas: []
      })
    ).rejects.toThrow();

    // ADMIN envía a Depósito y registra cobro -> VENTAS y DEPOSITO quedan bloqueados (403 AuthError)
    mockSession = {
      userId: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      tokenVersion: 1
    };
    await enviarPedidosACuentasDal({
      pedidoIds: [pedVentas.id],
      usuarioIds: [DEP_ID],
      prioridad: 'NORMAL',
      notaEnvio: '',
      forzarSinPago: true
    });

    mockSession = {
      userId: DEP_ID,
      username: 'deposito',
      nombre: 'Depósito',
      rol: 'DEPOSITO',
      tokenVersion: 1
    };
    await expect(
      actualizarArticulosPedidoDal({
        pedidoId: pedVentas.id,
        expectedVersion: okVentas.pedido.version,
        lineas: [resolverLinea(dino, 1, '')]
      })
    ).rejects.toBeInstanceOf(AuthError);

    mockSession = {
      userId: JUANCHO_ID,
      username: 'juancho',
      nombre: 'Juancho',
      rol: 'VENTAS',
      tokenVersion: 1
    };
    await expect(
      actualizarArticulosPedidoDal({
        pedidoId: pedVentas.id,
        expectedVersion: okVentas.pedido.version,
        lineas: [resolverLinea(dino, 1, '')]
      })
    ).rejects.toBeInstanceOf(AuthError);
  });

  it('55. Test G: generarMensajeAjusteStockWhatsApp formatea variaciones, total y saldo', () => {
    const msgPendiente = generarMensajeAjusteStockWhatsApp(
      {
        cliente: 'Casa Lorenzo',
        numeroPedido: '0042',
        totalFacturado: 3390120.4,
        saldoPendiente: 1200000
      },
      {
        cambios: [
          {
            tipo: 'modificado',
            nombreProducto: 'Chaleco Shark',
            detalle: 'T1 1-3 años',
            cantidadAnterior: 3,
            cantidadNueva: 1
          },
          {
            tipo: 'quitado',
            nombreProducto: 'Antiparras Dino',
            cantidadAnterior: 3,
            cantidadNueva: 0
          },
          {
            tipo: 'agregado',
            nombreProducto: 'Piluso Shark',
            cantidadAnterior: 0,
            cantidadNueva: 2
          }
        ],
        totalAnterior: 3560673.05,
        totalNuevo: 3390120.4,
        unidadesAnteriores: 8,
        unidadesNuevas: 6,
        saldoAFavorGenerado: 0,
        diferenciaCheque: 0,
        avisoRelecturaDeposito: false,
        incidenciaResuelta: false
      }
    );

    expect(msgPendiente).toContain('Casa Lorenzo');
    expect(msgPendiente).toContain('#0042');
    expect(msgPendiente).toContain('Chaleco Shark (T1 1-3 años): de 3 a 1 u.');
    expect(msgPendiente).toContain('Antiparras Dino: quitado por falta de stock');
    expect(msgPendiente).toContain('Piluso Shark: agregado (2 u.)');
    expect(msgPendiente).toContain('Saldo pendiente:');

    const msgSaldoFavor = generarMensajeAjusteStockWhatsApp(
      {
        cliente: 'Casa Lorenzo',
        numeroPedido: '0042',
        totalFacturado: 850000,
        saldoPendiente: 0
      },
      {
        cambios: [
          {
            tipo: 'modificado',
            nombreProducto: 'Antiparras Dino',
            cantidadAnterior: 10,
            cantidadNueva: 8
          }
        ],
        totalAnterior: 1000000,
        totalNuevo: 850000,
        unidadesAnteriores: 10,
        unidadesNuevas: 8,
        saldoAFavorGenerado: 150000,
        diferenciaCheque: 0,
        avisoRelecturaDeposito: false,
        incidenciaResuelta: false
      }
    );
    expect(msgSaldoFavor).toContain('Saldo a favor:');
  });
});
