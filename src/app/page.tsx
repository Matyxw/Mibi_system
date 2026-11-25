import React from 'react';
import { redirect } from 'next/navigation';
import { AuthError, requireSession } from '@/lib/dal/auth';
import { listarPedidosDal } from '@/lib/dal/pedidos';
import { isLocalFallbackMode } from '@/lib/db';
import { ErpShell } from '@/components/erp-shell';

export const dynamic = 'force-dynamic';

export default async function HomePage(): Promise<React.JSX.Element> {
  let session;
  try {
    session = await requireSession();
  } catch (err) {
    if (err instanceof AuthError) {
      redirect('/login?revoked=1');
    }
    throw err;
  }

  const initialPedidos = await listarPedidosDal();
  return (
    <ErpShell
      session={session}
      initialPedidos={initialPedidos}
      isLocalDev={isLocalFallbackMode()}
    />
  );
}
