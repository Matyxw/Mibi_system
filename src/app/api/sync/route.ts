import { NextResponse } from 'next/server';
import { AuthError } from '@/lib/dal/auth';
import { getSyncStatusDal } from '@/lib/dal/pedidos';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

export async function GET(): Promise<NextResponse> {
  try {
    const status = await getSyncStatusDal();
    return NextResponse.json(status, {
      headers: {
        'Cache-Control': 'no-store'
      }
    });
  } catch (err) {
    if (err instanceof AuthError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: 'Error al consultar sincronización' }, { status: 500 });
  }
}
