import React from 'react';
import { redirect } from 'next/navigation';
import { getSessionFromCookies } from '@/lib/auth/session';
import { isLocalFallbackMode } from '@/lib/db';
import { verifyActiveSessionUser } from '@/lib/dal/auth';
import { LoginForm } from '@/components/login-form';

export const dynamic = 'force-dynamic';

export default async function LoginPage(): Promise<React.JSX.Element> {
  const session = await getSessionFromCookies();
  if (session) {
    const active = await verifyActiveSessionUser(session);
    if (active) {
      redirect('/');
    }
  }

  return <LoginForm isLocalDev={isLocalFallbackMode()} />;
}
