import React from 'react';
import { redirect } from 'next/navigation';
import { AuthError, requireSession } from '@/lib/dal/auth';
import { listarProductosCustomDal } from '@/lib/dal/catalogo';
import { listarPedidosDal } from '@/lib/dal/pedidos';
import { listarCuentasAsignablesDal } from '@/lib/dal/usuarios';
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

  const [initialPedidos, initialCatalogoCustom, cuentasAsignables] = await Promise.all([
    listarPedidosDal(),
    listarProductosCustomDal(false),
    session.rol === 'ADMIN' ? listarCuentasAsignablesDal() : Promise.resolve([])
  ]);

  return (
    <ErpShell
      session={session}
      initialPedidos={initialPedidos}
      initialCatalogoCustom={initialCatalogoCustom}
      initialCuentasAsignables={cuentasAsignables}
      isLocalDev={isLocalFallbackMode()}
    />
  );
}
