import 'server-only';

import { eq } from 'drizzle-orm';
import { getDb, isLocalFallbackMode, readLocalStore } from '@/lib/db';
import { usuarios } from '@/lib/db/schema';
import { getSessionFromCookies } from '@/lib/auth/session';
import type { SessionUser } from '@/lib/types';

export class AuthError extends Error {
  public readonly status: number;

  constructor(message: string, status = 401) {
    super(message);
    this.name = 'AuthError';
    this.status = status;
  }
}

export const verifyActiveSessionUser = async (session: SessionUser): Promise<SessionUser | null> => {
  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const u = store.usuarios.find((x) => x.id === session.userId);
    if (!u || !u.activo || u.tokenVersion !== session.tokenVersion) {
      return null;
    }
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
    .select({
      id: usuarios.id,
      username: usuarios.username,
      nombre: usuarios.nombre,
      rol: usuarios.rol,
      activo: usuarios.activo,
      tokenVersion: usuarios.tokenVersion
    })
    .from(usuarios)
    .where(eq(usuarios.id, session.userId))
    .limit(1);

  const u = rows[0];
  if (!u || !u.activo || u.tokenVersion !== session.tokenVersion) {
    return null;
  }
  return {
    userId: u.id,
    username: u.username,
    nombre: u.nombre,
    rol: u.rol,
    tokenVersion: u.tokenVersion
  };
};

export const requireSession = async (): Promise<SessionUser> => {
  const session = await getSessionFromCookies();
  if (!session) {
    throw new AuthError('Sesión inválida o expirada.', 401);
  }
  const verified = await verifyActiveSessionUser(session);
  if (!verified) {
    throw new AuthError('Sesión revocada o usuario deshabilitado.', 401);
  }
  return verified;
};

export const requireAdmin = async (): Promise<SessionUser> => {
  const session = await requireSession();
  if (session.rol !== 'ADMIN') {
    throw new AuthError('Acción exclusiva para Administración.', 403);
  }
  return session;
};
