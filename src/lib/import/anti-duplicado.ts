import { normalizarTexto } from '@/lib/excel-parser';
import type { PedidoBase } from '@/lib/types';

const SIETE_DIAS_MS = 7 * 24 * 60 * 60 * 1000;

export const extraerNumeroFacturaAfip = (observaciones: string): string | null => {
  const m = observaciones.match(/Factura\s+AFIP\s+(\d{5}-\d{8})/i);
  return m?.[1] ?? null;
};

export const detectarPedidoDuplicado = <
  T extends Pick<PedidoBase, 'id' | 'numeroPedido' | 'cliente' | 'observaciones' | 'createdAt' | 'deletedAt' | 'lineas'>
>(params: {
  cliente: string;
  unidades: number;
  observaciones: string;
  ahora: number;
  pedidos: ReadonlyArray<T>;
}): T | null => {
  const { cliente, unidades, observaciones, ahora, pedidos } = params;
  const clienteNorm = normalizarTexto(cliente);
  const facturaNueva = extraerNumeroFacturaAfip(observaciones);

  for (const p of pedidos) {
    if (p.deletedAt) continue;

    if (facturaNueva) {
      const facturaExistente = extraerNumeroFacturaAfip(p.observaciones);
      if (facturaExistente && facturaExistente === facturaNueva) {
        return p;
      }
    }

    if (!clienteNorm || unidades <= 0) continue;
    const pClienteNorm = normalizarTexto(p.cliente);
    if (pClienteNorm !== clienteNorm) continue;

    const deltaMs = ahora - p.createdAt;
    if (deltaMs < -86400000 || deltaMs > SIETE_DIAS_MS) continue;

    const unidadesExistente = p.lineas.reduce((sum, l) => sum + l.cantidadTotal, 0);
    if (unidadesExistente === unidades) {
      return p;
    }
  }

  return null;
};
