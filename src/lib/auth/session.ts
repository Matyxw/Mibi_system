import 'server-only';

import { cookies } from 'next/headers';
import type { SessionUser } from '@/lib/types';
import {
  getSessionCookieName,
  SESSION_DURATION_SECONDS,
  signSessionToken,
  verifySessionToken
} from './jwt';

export const setSessionCookie = async (user: SessionUser): Promise<void> => {
  const token = await signSessionToken(user);
  const cookieStore = await cookies();
  cookieStore.set(getSessionCookieName(), token, {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'lax',
    path: '/',
    maxAge: SESSION_DURATION_SECONDS
  });
};

export const clearSessionCookie = async (): Promise<void> => {
  const cookieStore = await cookies();
  cookieStore.set(getSessionCookieName(), '', {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'lax',
    path: '/',
    maxAge: 0
  });
};

export const getSessionFromCookies = async (): Promise<SessionUser | null> => {
  const cookieStore = await cookies();
  const token = cookieStore.get(getSessionCookieName())?.value;
  if (!token) return null;
  return verifySessionToken(token);
};
