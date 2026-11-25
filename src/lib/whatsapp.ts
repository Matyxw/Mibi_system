import { CATALOGO_BY_GRUPO } from './catalogo-mibi';
import type { LineaPedido, Pedido, ProductoCatalogo } from './types';

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
