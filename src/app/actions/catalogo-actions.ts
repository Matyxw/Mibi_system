'use server';

import { revalidatePath } from 'next/cache';
import { z } from 'zod';
import { AuthError, requireAdmin, requireSession } from '@/lib/dal/auth';
import {
  crearProductoCatalogoDal,
  editarProductoCatalogoDal,
  listarProductosCustomDal,
  toggleArchivoProductoCatalogoDal
} from '@/lib/dal/catalogo';
import type { ProductoCustomRecord } from '@/lib/types';

const VarianteInputSchema = z.object({
  codigo: z.string().trim().max(40).optional(),
  etiqueta: z.string().trim().min(1, 'Ingresá el nombre del talle o color.').max(100),
  aliases: z.array(z.string().trim().max(60)).optional()
});

const ProductoInputSchema = z.object({
  nombre: z.string().trim().min(1, 'Ingresá el nombre del producto.').max(160),
  rubro: z.string().trim().min(1, 'Ingresá el rubro.').max(60),
  precioNeto: z.number().min(0, 'El precio neto no puede ser negativo.').max(100_000_000),
  variantes: z.array(VarianteInputSchema).min(1, 'Agregá al menos una variante.')
});

const ArchivarSchema = z.object({
  id: z.string().min(1),
  activo: z.boolean()
});

export interface CatalogoActionResponse<T> {
  ok: boolean;
  data?: T;
  error?: string;
}

const toCatalogoError = (err: unknown, fallback: string): string => {
  if (err instanceof AuthError) return err.message;
  if (err instanceof Error && err.message) return err.message;
  return fallback;
};

export async function listarProductosCustomAction(
  includeArchived = false
): Promise<CatalogoActionResponse<ProductoCustomRecord[]>> {
  try {
    await requireSession();
    const list = await listarProductosCustomDal(includeArchived);
    return { ok: true, data: list };
  } catch (err) {
    return { ok: false, error: toCatalogoError(err, 'No se pudo obtener el catálogo.') };
  }
};

export async function crearProductoCatalogoAction(
  raw: z.infer<typeof ProductoInputSchema>
): Promise<CatalogoActionResponse<ProductoCustomRecord>> {
  try {
    await requireSession();
    const parsed = ProductoInputSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: parsed.error.issues[0]?.message ?? 'Datos de producto inválidos.' };
    }
    const created = await crearProductoCatalogoDal(parsed.data);
    revalidatePath('/');
    return { ok: true, data: created };
  } catch (err) {
    return { ok: false, error: toCatalogoError(err, 'No se pudo crear el producto.') };
  }
}

export async function editarProductoCatalogoAction(
  id: string,
  raw: z.infer<typeof ProductoInputSchema>
): Promise<CatalogoActionResponse<ProductoCustomRecord>> {
  try {
    await requireAdmin();
    const parsed = ProductoInputSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: parsed.error.issues[0]?.message ?? 'Datos de producto inválidos.' };
    }
    const updated = await editarProductoCatalogoDal(id, parsed.data);
    revalidatePath('/');
    return { ok: true, data: updated };
  } catch (err) {
    return { ok: false, error: toCatalogoError(err, 'No se pudo actualizar el producto.') };
  }
}

export async function archivarProductoCatalogoAction(
  raw: z.infer<typeof ArchivarSchema>
): Promise<CatalogoActionResponse<ProductoCustomRecord>> {
  try {
    await requireAdmin();
    const parsed = ArchivarSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: 'Datos inválidos.' };
    }
    const updated = await toggleArchivoProductoCatalogoDal(parsed.data.id, parsed.data.activo);
    revalidatePath('/');
    return { ok: true, data: updated };
  } catch (err) {
    return { ok: false, error: toCatalogoError(err, 'No se pudo actualizar el estado del producto.') };
  }
}
