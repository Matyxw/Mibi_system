import { NextResponse, type NextRequest } from 'next/server';
import { AuthError, requireSession } from '@/lib/dal/auth';
import { obtenerPedidoPorIdDal } from '@/lib/dal/pedidos';
import { generarBufferHojaDeposito } from '@/lib/excel-exporter';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

export async function GET(
  _req: NextRequest,
  context: { params: Promise<{ id: string }> }
): Promise<NextResponse> {
  try {
    await requireSession();
    const { id } = await context.params;
    const pedido = await obtenerPedidoPorIdDal(id);
    if (!pedido) {
      return NextResponse.json({ error: 'Pedido no encontrado.' }, { status: 404 });
    }

    const res = generarBufferHojaDeposito(pedido);
    if (!res.success || !res.buffer || !res.filename) {
      return NextResponse.json(
        { error: res.errors[0] ?? 'No se puede exportar hoja de depósito.' },
        { status: 409 }
      );
    }

    const bodyBuffer = res.buffer.buffer.slice(
      res.buffer.byteOffset,
      res.buffer.byteOffset + res.buffer.byteLength
    ) as ArrayBuffer;

    return new NextResponse(bodyBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${res.filename}"`,
        'Cache-Control': 'no-store'
      }
    });
  } catch (err) {
    if (err instanceof AuthError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: 'Error al generar Hoja Depósito.' }, { status: 500 });
  }
}
