import { CATALOGO_BY_GRUPO } from './catalogo-mibi';
import { formatARS } from './finanzas';
import type { LineaPedido, Pedido, ProductoCatalogo, ResumenAjusteArticulos } from './types';

export const generarMensajeWhatsApp = (
  cliente: string,
  numeroPedido: string,
  lineas: ReadonlyArray<LineaPedido>,
  catalogoExtra: ReadonlyArray<ProductoCatalogo> = []
): string => {
  const pendientes = lineas.filter((l) => l.requiereRevision || l.estado !== 'OK');
  if (pendientes.length === 0) return '';

  const extraByGrupo = new Map<string, ProductoCatalogo>();
  for (const p of catalogoExtra) {
    extraByGrupo.set(p.idGrupo, p);
  }

  const lista = pendientes
    .map((l) => {
      const prod = extraByGrupo.get(l.idGrupo) ?? CATALOGO_BY_GRUPO.get(l.idGrupo);
      const variantes = prod ? prod.variantes : l.desglose;
      const opciones = variantes.map((v) => `${v.etiqueta} (${v.codigo})`).join(' / ');
      return `• ${l.nombreProducto} — ${l.cantidadTotal} u. (elegir entre: ${opciones})`;
    })
    .join('\n');

  const saludo = cliente ? `¡Hola ${cliente}!` : '¡Hola!';
  const ref = numeroPedido ? ` por tu pedido #${numeroPedido}` : ' por tu pedido';
  return `${saludo} Te escribimos de Mibi®${ref}.\n\nPara pasarlo a depósito sin errores, nos falta confirmar talle o color en:\n\n${lista}\n\nApenas nos confirmes esas cantidades lo dejamos listo para armar. ¡Muchas gracias!`;
};

export const generarMensajeDespachoWhatsApp = (
  pedido: Pick<Pedido, 'cliente' | 'numeroPedido' | 'transporte' | 'nroSeguimiento'>
): string => {
  const transporte = pedido.transporte.trim();
  const guia = pedido.nroSeguimiento.trim();
  if (!transporte || !guia) return '';
  return `¡Hola ${pedido.cliente}! Tu pedido #${pedido.numeroPedido} de Mibi® ya fue despachado por ${transporte}. Guía de seguimiento: ${guia}`;
};

export const generarMensajeAjusteStockWhatsApp = (
  pedido: Pick<Pedido, 'cliente' | 'numeroPedido' | 'totalFacturado' | 'saldoPendiente'>,
  resumenAjuste: ResumenAjusteArticulos
): string => {
  const saludo = pedido.cliente.trim() ? `¡Hola ${pedido.cliente.trim()}!` : '¡Hola!';
  const bullets = resumenAjuste.cambios
    .map((c) => {
      const detalle = c.detalle ? ` (${c.detalle})` : '';
      if (c.tipo === 'quitado') {
        return `• ${c.nombreProducto}${detalle}: quitado por falta de stock`;
      }
      if (c.tipo === 'agregado') {
        return `• ${c.nombreProducto}${detalle}: agregado (${c.cantidadNueva} u.)`;
      }
      return `• ${c.nombreProducto}${detalle}: de ${c.cantidadAnterior} a ${c.cantidadNueva} u.`;
    })
    .join('\n');

  const partesSaldo: string[] = [];
  if (resumenAjuste.saldoAFavorGenerado > 0.005) {
    partesSaldo.push(`Saldo a favor: ${formatARS(resumenAjuste.saldoAFavorGenerado)}`);
  } else if (pedido.saldoPendiente > 0.005) {
    partesSaldo.push(`Saldo pendiente: ${formatARS(pedido.saldoPendiente)}`);
  }
  const suffixSaldo = partesSaldo.length > 0 ? ` (${partesSaldo.join(' · ')})` : '';
  const cuerpoLista = bullets ? `\n\n${bullets}` : '';

  return `${saludo} Te escribimos de Mibi® por tu pedido #${pedido.numeroPedido}. Te avisamos que ajustamos las cantidades según disponibilidad de stock:${cuerpoLista}\n\nEl nuevo total actualizado quedó en ${formatARS(resumenAjuste.totalNuevo)}${suffixSaldo}. ¿Te parece bien que avancemos así? ¡Muchas gracias!`;
};

export const buildWhatsAppUrl = (telefonoRaw: string, mensaje: string): string | null => {
  const texto = mensaje.trim();
  if (!texto) return null;

  let digits = telefonoRaw.replace(/\D+/g, '');
  if (!digits) return null;

  if (digits.startsWith('00')) {
    digits = digits.slice(2);
  } else if (digits.startsWith('0')) {
    digits = digits.slice(1);
  }

  if (digits.length === 10) {
    digits = `549${digits}`;
  } else if (digits.length === 12 && digits.startsWith('54') && !digits.startsWith('549')) {
    digits = `549${digits.slice(2)}`;
  }

  if (digits.length < 10) return null;
  return `https://wa.me/${digits}?text=${encodeURIComponent(texto)}`;
};
