import { NextResponse, type NextRequest } from 'next/server';
import { z } from 'zod';
import { AuthError, requireSession } from '@/lib/dal/auth';
import { listarPedidosDal, obtenerPedidosPorIdsDal } from '@/lib/dal/pedidos';
import { generarBufferTableroMaestro } from '@/lib/excel-exporter';
import { calcularResumenFinanciero } from '@/lib/finanzas';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

const BodySchema = z.object({
  ids: z.array(z.string().min(1)).optional(),
  sufijo: z.string().max(40).optional()
});

export async function POST(req: NextRequest): Promise<NextResponse> {
  try {
    await requireSession();
    const json = (await req.json().catch(() => ({}))) as unknown;
    const parsed = BodySchema.safeParse(json);
    if (!parsed.success) {
      return NextResponse.json({ error: 'Parámetros inválidos.' }, { status: 400 });
    }

    const ids = parsed.data.ids;
    const pedidos = ids && ids.length > 0 ? await obtenerPedidosPorIdsDal(ids) : await listarPedidosDal();
    const resumen = calcularResumenFinanciero(pedidos);
    const sufijo = parsed.data.sufijo ?? (ids && ids.length > 0 ? `SELECCION_${pedidos.length}` : '');

    const { buffer, filename } = await generarBufferTableroMaestro(pedidos, resumen, sufijo);
    const bodyBuffer = buffer.buffer.slice(
      buffer.byteOffset,
      buffer.byteOffset + buffer.byteLength
    ) as ArrayBuffer;

    return new NextResponse(bodyBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${filename}"`,
        'Cache-Control': 'no-store'
      }
    });
  } catch (err) {
    if (err instanceof AuthError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: 'Error al exportar Tablero Maestro.' }, { status: 500 });
  }
}
