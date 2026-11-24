import { SignJWT, jwtVerify } from 'jose';
import type { RolUsuario, SessionUser } from '@/lib/types';

export const SESSION_DURATION_SECONDS = 60 * 60 * 12; // 12 horas

export const getSessionCookieName = (): string =>
  process.env.NODE_ENV === 'production' ? '__Host-mibi_session' : 'mibi_session';

const getJwtSecretKey = (): Uint8Array => {
  const raw = process.env.AUTH_SECRET;
  if (process.env.NODE_ENV === 'production') {
    if (!raw || raw.length < 32) {
      throw new Error('FATAL: AUTH_SECRET debe configurarse con al menos 32 bytes en producción.');
    }
    return new TextEncoder().encode(raw);
  }
  const devSecret =
    raw && raw.length >= 32 ? raw : 'mibi-dev-only-local-bunker-secret-key-min-32-bytes';
  return new TextEncoder().encode(devSecret);
};

export const signSessionToken = async (user: SessionUser): Promise<string> => {
  return new SignJWT({
    userId: user.userId,
    username: user.username,
    nombre: user.nombre,
    rol: user.rol,
    tokenVersion: user.tokenVersion
  })
    .setProtectedHeader({ alg: 'HS256' })
    .setSubject(user.userId)
    .setIssuedAt()
    .setExpirationTime(`${SESSION_DURATION_SECONDS}s`)
    .sign(getJwtSecretKey());
};

export const verifySessionToken = async (token: string): Promise<SessionUser | null> => {
  if (!token) return null;
  try {
    const { payload } = await jwtVerify(token, getJwtSecretKey(), {
      algorithms: ['HS256']
    });
    const userId = typeof payload['userId'] === 'string' ? payload['userId'] : '';
    const username = typeof payload['username'] === 'string' ? payload['username'] : '';
    const nombre = typeof payload['nombre'] === 'string' ? payload['nombre'] : '';
    const rolRaw = payload['rol'];
    const rol: RolUsuario | null =
      rolRaw === 'ADMIN' || rolRaw === 'VENTAS_DEPOSITO' ? rolRaw : null;
    const tokenVersion = typeof payload['tokenVersion'] === 'number' ? payload['tokenVersion'] : 0;

    if (!userId || !username || !nombre || !rol || tokenVersion < 1) {
      return null;
    }
    return { userId, username, nombre, rol, tokenVersion };
  } catch {
    return null;
  }
};
