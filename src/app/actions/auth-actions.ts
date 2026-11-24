'use server';

import { headers } from 'next/headers';
import { redirect } from 'next/navigation';
import { z } from 'zod';
import { clearLoginRateLimit, checkLoginRateLimit, recordLoginFailure } from '@/lib/auth/rate-limit';
import { clearSessionCookie, setSessionCookie } from '@/lib/auth/session';
import { authenticateUser } from '@/lib/dal/usuarios';

const LoginSchema = z.object({
  username: z.string().trim().min(1).max(64),
  password: z.string().min(1).max(128)
});

export interface LoginActionResult {
  ok: boolean;
  error?: string;
}

export async function loginAction(formData: {
  username: string;
  password: string;
}): Promise<LoginActionResult> {
  const parsed = LoginSchema.safeParse(formData);
  if (!parsed.success) {
    return { ok: false, error: 'Usuario o contraseña incorrectos' };
  }

  const hdrs = await headers();
  const forwarded = hdrs.get('x-forwarded-for') ?? '';
  const ip = forwarded.split(',')[0]?.trim() || hdrs.get('x-real-ip') || 'unknown';
  const username = parsed.data.username.toLowerCase();

  const rl = checkLoginRateLimit(ip, username);
  if (!rl.allowed) {
    return {
      ok: false,
      error: `Demasiados intentos fallidos. Esperá ${Math.ceil(rl.retryAfterSec / 60)} min.`
    };
  }

  const user = await authenticateUser(username, parsed.data.password);
  if (!user) {
    recordLoginFailure(ip, username);
    return { ok: false, error: 'Usuario o contraseña incorrectos' };
  }

  clearLoginRateLimit(ip, username);
  await setSessionCookie(user);
  return { ok: true };
}

export async function logoutAction(): Promise<void> {
  await clearSessionCookie();
  redirect('/login');
}
