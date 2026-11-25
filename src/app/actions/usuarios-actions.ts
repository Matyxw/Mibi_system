'use server';

import { revalidatePath } from 'next/cache';
import { z } from 'zod';
import { AuthError, requireAdmin } from '@/lib/dal/auth';
import {
  cambiarRolUsuarioDal,
  crearUsuarioDal,
  listarCuentasAsignablesDal,
  listarUsuariosDal,
  rotarPasswordUsuarioDal,
  toggleActivoUsuarioDal
} from '@/lib/dal/usuarios';
import type { UsuarioEquipo } from '@/lib/types';

const CrearUsuarioSchema = z.object({
  username: z
    .string()
    .trim()
    .min(2)
    .max(40)
    .regex(/^[a-zA-Z0-9._-]+$/, 'Usá solo letras, números, punto o guion.'),
  nombre: z.string().trim().min(2).max(80),
  rol: z.enum(['ADMIN', 'VENTAS', 'DEPOSITO', 'VENTAS_DEPOSITO']),
  password: z.string().min(4).max(128)
});

const RotarPasswordSchema = z.object({
  userId: z.string().min(1),
  nuevaPassword: z.string().min(4).max(128)
});

const ToggleActivoSchema = z.object({
  userId: z.string().min(1),
  activo: z.boolean()
});

const CambiarRolSchema = z.object({
  userId: z.string().min(1),
  rol: z.enum(['ADMIN', 'VENTAS', 'DEPOSITO', 'VENTAS_DEPOSITO'])
});

export interface UsuarioActionResponse<T> {
  ok: boolean;
  data?: T;
  error?: string;
}

const toErr = (err: unknown, fallback: string): string => {
  if (err instanceof AuthError) return err.message;
  if (err instanceof Error && err.message) return err.message;
  return fallback;
};

export async function listarUsuariosAction(): Promise<UsuarioActionResponse<UsuarioEquipo[]>> {
  try {
    await requireAdmin();
    const list = await listarUsuariosDal();
    return { ok: true, data: list };
  } catch (err) {
    return { ok: false, error: toErr(err, 'No se pudo listar el equipo.') };
  }
}

export async function crearUsuarioAction(
  raw: z.infer<typeof CrearUsuarioSchema>
): Promise<UsuarioActionResponse<UsuarioEquipo>> {
  try {
    await requireAdmin();
    const parsed = CrearUsuarioSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: parsed.error.issues[0]?.message ?? 'Datos de usuario inválidos.' };
    }
    const created = await crearUsuarioDal(parsed.data);
    revalidatePath('/');
    return { ok: true, data: created };
  } catch (err) {
    return { ok: false, error: toErr(err, 'No se pudo crear el usuario.') };
  }
}

export async function rotarPasswordAction(
  raw: z.infer<typeof RotarPasswordSchema>
): Promise<UsuarioActionResponse<{ tokenVersion: number }>> {
  try {
    await requireAdmin();
    const parsed = RotarPasswordSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: 'La contraseña debe tener al menos 4 caracteres.' };
    }
    const res = await rotarPasswordUsuarioDal(parsed.data.userId, parsed.data.nuevaPassword);
    revalidatePath('/');
    return { ok: true, data: res };
  } catch (err) {
    return { ok: false, error: toErr(err, 'No se pudo actualizar la contraseña.') };
  }
}

export async function toggleActivoUsuarioAction(
  raw: z.infer<typeof ToggleActivoSchema>
): Promise<UsuarioActionResponse<{ ok: true }>> {
  try {
    await requireAdmin();
    const parsed = ToggleActivoSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: 'Datos inválidos.' };
    }
    await toggleActivoUsuarioDal(parsed.data.userId, parsed.data.activo);
    revalidatePath('/');
    return { ok: true, data: { ok: true } };
  } catch (err) {
    return { ok: false, error: toErr(err, 'No se pudo actualizar el estado del usuario.') };
  }
}

export async function cambiarRolUsuarioAction(
  raw: z.infer<typeof CambiarRolSchema>
): Promise<UsuarioActionResponse<{ ok: true }>> {
  try {
    await requireAdmin();
    const parsed = CambiarRolSchema.safeParse(raw);
    if (!parsed.success) {
      return { ok: false, error: 'Datos de rol inválidos.' };
    }
    await cambiarRolUsuarioDal(parsed.data.userId, parsed.data.rol);
    revalidatePath('/');
    return { ok: true, data: { ok: true } };
  } catch (err) {
    return { ok: false, error: toErr(err, 'No se pudo cambiar el rol del usuario.') };
  }
}

export async function listarCuentasAsignablesAction(): Promise<UsuarioActionResponse<UsuarioEquipo[]>> {
  try {
    await requireAdmin();
    const list = await listarCuentasAsignablesDal();
    return { ok: true, data: list };
  } catch (err) {
    return { ok: false, error: toErr(err, 'No se pudieron obtener las cuentas.') };
  }
}
