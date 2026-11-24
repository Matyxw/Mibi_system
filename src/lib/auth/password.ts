import 'server-only';

import { randomBytes, scryptSync, timingSafeEqual } from 'node:crypto';

const SCRYPT_KEYLEN = 64;
const SCRYPT_PREFIX = 'scrypt$v1';

export const hashPassword = (plainPassword: string): string => {
  const normalized = plainPassword.trim();
  if (normalized.length < 4) {
    throw new Error('La contraseña debe tener al menos 4 caracteres.');
  }
  const salt = randomBytes(16).toString('hex');
  const derived = scryptSync(normalized, salt, SCRYPT_KEYLEN).toString('hex');
  return `${SCRYPT_PREFIX}$${salt}$${derived}`;
};

export const verifyPassword = (plainPassword: string, storedHash: string): boolean => {
  const parts = storedHash.split('$');
  if (parts.length !== 4 || parts[0] !== 'scrypt' || parts[1] !== 'v1') {
    return false;
  }
  const salt = parts[2];
  const expectedHex = parts[3];
  if (!salt || !expectedHex) return false;

  const expectedBuf = Buffer.from(expectedHex, 'hex');
  const derivedBuf = scryptSync(plainPassword.trim(), salt, expectedBuf.length);
  if (derivedBuf.length !== expectedBuf.length) return false;
  return timingSafeEqual(derivedBuf, expectedBuf);
};
