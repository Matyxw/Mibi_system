'use server';

import { revalidatePath } from 'next/cache';
import { z } from 'zod';
import { AuthError, requireAdmin, requireSession } from '@/lib/dal/auth';
import {
  batchUpdatePedidosDal,
  crearPedidoDal,
  importarPedidosHistoricosDal,
  listarPedidosDal,
  patchPedidoDal,
  reiniciarSistemaDal,
  restorePedidosDal,
  softDeletePedidosDal,
  type BatchActionInput,
  type PatchPedidoInput
} from '@/lib/dal/pedidos';
import { importarPlanillaGeneral } from '@/lib/excel-parser';
import type { Pedido } from '@/lib/types';

const DesgloseSchema = z.object({
  codigo: z.string().min(1),
  etiqueta: z.string(),
  cantidad: z.number().int().min(0)
});

const LineaSchema = z.object({
  idLinea: z.string().min(1),
  idGrupo: z.string().min(1),
  nombreProducto: z.string().min(1),
  rubro: z.string(),
  precioUnitarioNeto: z.number().min(0),
  cantidadTotal: z.number().int().min(1),
  notaCliente: z.string(),
  estado: z.enum(['OK', 'REQUIERE_REVISION']),
  requiereRevision: z.boolean(),
  motivoRevision: z.string().optional(),
  desglose: z.array(DesgloseSchema)
});

const CrearPedidoSchema = z.object({
  cliente: z.string().trim().min(1, 'Ingresá el nombre del cliente o negocio.').max(200),
  vendedor: z.string().trim().min(1).max(80),
  whatsapp: z.string().trim().max(80),
  fechaPedido: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  despachoPrevisto: z.string().max(20),
  observaciones: z.string().max(1000),
  descuentoPorc: z.number().min(0).max(100),
  medioPago: z.enum(['Transferencia / Efectivo', 'Cheque', 'Sin pagar']),
  montoAbonado: z.number().min(0),
  montoCheque: z.number().min(0),
  fechaRecepcionCheque: z.string().max(20),
  fechaAcreditacionCheque: z.string().max(20),
  detalleCheque: z.string().max(200),
  lineas: z.array(LineaSchema).min(1, 'Cargá al menos un artículo.')
});

const PatchPedidoSchema = z.object({
  pedidoId: z.string().min(1),
  motivo: z.string().trim().min(1).max(300),
  patch: z.object({
    cliente: z.string().trim().min(1).max(200).optional(),
    vendedor: z.string().trim().min(1).max(80).optional(),
    whatsapp: z.string().trim().max(80).optional(),
    despachoPrevisto: z.string().max(20).optional(),
    descuentoPorc: z.number().min(0).max(100).optional(),
    medioPago: z.enum(['Transferencia / Efectivo', 'Cheque', 'Sin pagar']).optional(),
    montoAbonado: z.number().min(0).optional(),
    montoCheque: z.number().min(0).optional(),
    chequeAcreditado: z.boolean().optional(),
    fechaRecepcionCheque: z.string().max(20).optional(),
    fechaAcreditacionCheque: z.string().max(20).optional(),
    detalleCheque: z.string().max(200).optional(),
    estadoArmado: z.enum(['Por armar', 'En armado', 'Listo', 'Falta talle o color']).optional(),
    estadoEnvio: z.enum(['Pendiente', 'Despachado', 'Entregado']).optional(),
    fechaDespacho: z.string().max(20).optional(),
    fechaEntrega: z.string().max(20).optional(),
    transporte: z.string().max(120).optional(),
    nroSeguimiento: z.string().max(120).optional(),
    observaciones: z.string().max(1000).optional(),
    lineas: z.array(LineaSchema).optional()
  })
});

const BatchActionSchema = z.object({
  ids: z.array(z.string().min(1)).min(1),
  action: z.discriminatedUnion('type', [
    z.object({ type: z.literal('armado'), estadoArmado: z.enum(['En armado', 'Listo']) }),
    z.object({
      type: z.literal('envio'),
      estadoEnvio: z.enum(['Despachado', 'Entregado']),
      transporte: z.string().max(120).optional()
    }),
    z.object({ type: z.literal('pago'), estado: z.enum(['Pagado', 'Falta pago']) }),
    z.object({ type: z.literal('vendedor'), vendedor: z.string().trim().min(1).max(80) })
  ])
});

const IdsSchema = z.object({
  ids: z.array(z.string().min(1)).min(1)
});

export interface ActionResponse<T> {
  ok: boolean;
  data?: T;
  error?: string;
}

const toActionError = (err: unknown, fallback: string): string => {
  if (err instanceof AuthError) return err.message;
  if (err instanceof Error && err.message) return err.message;
  return fallback;
};

export async function listarPedidosAction(): Promise<ActionResponse<Pedido[]>> {
  try {
    await requireSession();
    const data = await listarPedidosDal();
    return { ok: true, data };
  } catch (err) {
    return { ok: false, error: toActionError(err, 'No se pudieron obtener los pedidos.') };
  }
}

export async function crearPedidoAction(
  raw: z.infer<typeof CrearPedidoSchema>
): Promise<ActionResponse<Pedido>> {
  try {
    await requireSession();
    const parsed = CrearPedidoSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: parsed.error.issues[0]?.message ?? 'Datos del pedido inválidos.' };
    }
    const pedido = await crearPedidoDal(parsed.data);
    revalidatePath('/');
    return { ok: true, data: pedido };
  } catch (err) {
    return { ok: false, error: toActionError(err, 'No se pudo guardar el pedido.') };
  }
}

export async function patchPedidoAction(
  raw: z.infer<typeof PatchPedidoSchema>
): Promise<ActionResponse<Pedido>> {
  try {
    await requireSession();
    const parsed = PatchPedidoSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: parsed.error.issues[0]?.message ?? 'Datos inválidos al actualizar.' };
    }
    const patchInput: PatchPedidoInput = parsed.data.patch;
    const updated = await patchPedidoDal(parsed.data.pedidoId, patchInput, parsed.data.motivo);
    revalidatePath('/');
    return { ok: true, data: updated };
  } catch (err) {
    return { ok: false, error: toActionError(err, 'No se pudo actualizar el pedido.') };
  }
}

export async function batchPedidosAction(
  raw: z.infer<typeof BatchActionSchema>
): Promise<ActionResponse<{ actualizados: number; omitidos: string[] }>> {
  try {
    await requireSession();
    const parsed = BatchActionSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: 'Selección o acción en lote inválida.' };
    }
    const action: BatchActionInput = parsed.data.action;
    const res = await batchUpdatePedidosDal(parsed.data.ids, action);
    revalidatePath('/');
    return { ok: true, data: res };
  } catch (err) {
    return { ok: false, error: toActionError(err, 'No se pudo ejecutar la acción en lote.') };
  }
}

export async function softDeletePedidosAction(
  raw: z.infer<typeof IdsSchema>
): Promise<ActionResponse<{ eliminados: number }>> {
  try {
    await requireAdmin();
    const parsed = IdsSchema.safeParse(raw);
    if (!parsed.success) return { ok: false, error: 'IDs inválidos.' };
    const eliminados = await softDeletePedidosDal(parsed.data.ids);
    revalidatePath('/');
    return { ok: true, data: { eliminados } };
  } catch (err) {
    return { ok: false, error: toActionError(err, 'No se pudieron eliminar los pedidos.') };
  }
}

export async function restorePedidosAction(
  raw: z.infer<typeof IdsSchema>
): Promise<ActionResponse<{ restaurados: number }>> {
  try {
    await requireAdmin();
    const parsed = IdsSchema.safeParse(raw);
    if (!parsed.success) return { ok: false, error: 'IDs inválidos.' };
    const restaurados = await restorePedidosDal(parsed.data.ids);
    revalidatePath('/');
    return { ok: true, data: { restaurados } };
  } catch (err) {
    return { ok: false, error: toActionError(err, 'No se pudieron restaurar los pedidos.') };
  }
}

export async function importarPlanillaAction(
  formData: FormData
): Promise<ActionResponse<{ importados: number; duplicadosOmitidos: number }>> {
  try {
    await requireAdmin();
    const file = formData.get('file');
    if (!(file instanceof File)) {
      return { ok: false, error: 'No se recibió ningún archivo.' };
    }
    const buffer = await file.arrayBuffer();
    const existentes = await listarPedidosDal();
    const parsed = importarPlanillaGeneral(buffer, file.name, existentes);
    if (!parsed.ok) {
      return { ok: false, error: parsed.error ?? 'Error al leer la planilla.' };
    }
    if (parsed.nuevos.length === 0) {
      return {
        ok: true,
        data: { importados: 0, duplicadosOmitidos: parsed.duplicadosOmitidos }
      };
    }
    const res = await importarPedidosHistoricosDal(parsed.nuevos);
    revalidatePath('/');
    return {
      ok: true,
      data: {
        importados: res.importados,
        duplicadosOmitidos: parsed.duplicadosOmitidos + res.duplicadosOmitidos
      }
    };
  } catch (err) {
    return { ok: false, error: toActionError(err, 'Error al importar la planilla general.') };
  }
}

export async function reiniciarSistemaAction(): Promise<ActionResponse<{ ok: true }>> {
  try {
    await requireAdmin();
    await reiniciarSistemaDal();
    revalidatePath('/');
    return { ok: true, data: { ok: true } };
  } catch (err) {
    return { ok: false, error: toActionError(err, 'No se pudo reiniciar el sistema.') };
  }
}
