import 'server-only';

import { randomUUID } from 'node:crypto';
import { asc, eq } from 'drizzle-orm';
import { getDb, isLocalFallbackMode, readLocalStore, writeLocalStore } from '@/lib/db';
import { usuarios } from '@/lib/db/schema';
import { hashPassword, verifyPassword } from '@/lib/auth/password';
import type { RolUsuario, SessionUser, UsuarioEquipo } from '@/lib/types';
import { requireAdmin } from './auth';

export const authenticateUser = async (
  usernameRaw: string,
  passwordRaw: string
): Promise<SessionUser | null> => {
  const username = usernameRaw.trim().toLowerCase();
  const password = passwordRaw.trim();
  if (!username || !password) return null;

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const u = store.usuarios.find((x) => x.username.toLowerCase() === username);
    if (!u || !u.activo) return null;
    if (!verifyPassword(password, u.passwordHash)) return null;
    return {
      userId: u.id,
      username: u.username,
      nombre: u.nombre,
      rol: u.rol,
      tokenVersion: u.tokenVersion
    };
  }

  const db = getDb();
  const rows = await db
    .select()
    .from(usuarios)
    .where(eq(usuarios.username, username))
    .limit(1);
  const u = rows[0];
  if (!u || !u.activo) return null;
  if (!verifyPassword(password, u.passwordHash)) return null;
  return {
    userId: u.id,
    username: u.username,
    nombre: u.nombre,
    rol: u.rol,
    tokenVersion: u.tokenVersion
  };
};

export const listarUsuariosDal = async (): Promise<UsuarioEquipo[]> => {
  await requireAdmin();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    return store.usuarios.map((u) => ({
      id: u.id,
      username: u.username,
      nombre: u.nombre,
      rol: u.rol,
      activo: u.activo,
      tokenVersion: u.tokenVersion,
      createdAt: u.createdAt
    }));
  }

  const db = getDb();
  const rows = await db
    .select({
      id: usuarios.id,
      username: usuarios.username,
      nombre: usuarios.nombre,
      rol: usuarios.rol,
      activo: usuarios.activo,
      tokenVersion: usuarios.tokenVersion,
      createdAt: usuarios.createdAt
    })
    .from(usuarios)
    .orderBy(asc(usuarios.createdAt));

  return rows.map((r) => ({
    id: r.id,
    username: r.username,
    nombre: r.nombre,
    rol: r.rol,
    activo: r.activo,
    tokenVersion: r.tokenVersion,
    createdAt: r.createdAt.getTime()
  }));
};

export const crearUsuarioDal = async (input: {
  username: string;
  nombre: string;
  rol: RolUsuario;
  password: string;
}): Promise<UsuarioEquipo> => {
  await requireAdmin();
  const username = input.username.trim().toLowerCase();
  const nombre = input.nombre.trim();
  const passwordHash = hashPassword(input.password);

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    if (store.usuarios.some((u) => u.username.toLowerCase() === username)) {
      throw new Error('Ya existe un usuario con ese nombre de acceso.');
    }
    const now = Date.now();
    const nuevo = {
      id: randomUUID(),
      username,
      nombre,
      rol: input.rol,
      passwordHash,
      activo: true,
      tokenVersion: 1,
      createdAt: now,
      updatedAt: now
    };
    store.usuarios.push(nuevo);
    writeLocalStore(store);
    return {
      id: nuevo.id,
      username: nuevo.username,
      nombre: nuevo.nombre,
      rol: nuevo.rol,
      activo: nuevo.activo,
      tokenVersion: nuevo.tokenVersion,
      createdAt: nuevo.createdAt
    };
  }

  const db = getDb();
  const existing = await db
    .select({ id: usuarios.id })
    .from(usuarios)
    .where(eq(usuarios.username, username))
    .limit(1);
  if (existing.length > 0) {
    throw new Error('Ya existe un usuario con ese nombre de acceso.');
  }

  const [created] = await db
    .insert(usuarios)
    .values({
      username,
      nombre,
      rol: input.rol,
      passwordHash,
      activo: true,
      tokenVersion: 1
    })
    .returning();

  if (!created) {
    throw new Error('No se pudo crear el usuario.');
  }

  return {
    id: created.id,
    username: created.username,
    nombre: created.nombre,
    rol: created.rol,
    activo: created.activo,
    tokenVersion: created.tokenVersion,
    createdAt: created.createdAt.getTime()
  };
};

export const rotarPasswordUsuarioDal = async (
  targetUserId: string,
  nuevaPassword: string
): Promise<{ tokenVersion: number }> => {
  await requireAdmin();
  const passwordHash = hashPassword(nuevaPassword);

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const u = store.usuarios.find((x) => x.id === targetUserId);
    if (!u) throw new Error('Usuario no encontrado.');
    u.passwordHash = passwordHash;
    u.tokenVersion += 1;
    u.updatedAt = Date.now();
    writeLocalStore(store);
    return { tokenVersion: u.tokenVersion };
  }

  const db = getDb();
  const current = await db
    .select({ tokenVersion: usuarios.tokenVersion })
    .from(usuarios)
    .where(eq(usuarios.id, targetUserId))
    .limit(1);
  const curr = current[0];
  if (!curr) throw new Error('Usuario no encontrado.');

  const nextVersion = curr.tokenVersion + 1;
  await db
    .update(usuarios)
    .set({
      passwordHash,
      tokenVersion: nextVersion,
      updatedAt: new Date()
    })
    .where(eq(usuarios.id, targetUserId));

  return { tokenVersion: nextVersion };
};

export const toggleActivoUsuarioDal = async (
  targetUserId: string,
  activo: boolean
): Promise<void> => {
  const admin = await requireAdmin();
  if (admin.userId === targetUserId && !activo) {
    throw new Error('No podés desactivar tu propio usuario administrador.');
  }

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const u = store.usuarios.find((x) => x.id === targetUserId);
    if (!u) throw new Error('Usuario no encontrado.');
    u.activo = activo;
    u.tokenVersion += 1;
    u.updatedAt = Date.now();
    writeLocalStore(store);
    return;
  }

  const db = getDb();
  const current = await db
    .select({ tokenVersion: usuarios.tokenVersion })
    .from(usuarios)
    .where(eq(usuarios.id, targetUserId))
    .limit(1);
  const curr = current[0];
  if (!curr) throw new Error('Usuario no encontrado.');

  await db
    .update(usuarios)
    .set({
      activo,
      tokenVersion: curr.tokenVersion + 1,
      updatedAt: new Date()
    })
    .where(eq(usuarios.id, targetUserId));
};
