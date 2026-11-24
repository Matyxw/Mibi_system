import { NextResponse, type NextRequest } from 'next/server';
import { z } from 'zod';
import { AuthError, requireSession } from '@/lib/dal/auth';
import { obtenerPedidosPorIdsDal } from '@/lib/dal/pedidos';
import { generarBufferPickingConsolidado } from '@/lib/excel-exporter';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

const BodySchema = z.object({
  ids: z.array(z.string().min(1)).min(1)
});

export async function POST(req: NextRequest): Promise<NextResponse> {
  try {
    await requireSession();
    const json = (await req.json().catch(() => ({}))) as unknown;
    const parsed = BodySchema.safeParse(json);
    if (!parsed.success) {
      return NextResponse.json({ error: 'Seleccioná al menos un pedido.' }, { status: 400 });
    }

    const seleccionados = await obtenerPedidosPorIdsDal(parsed.data.ids);
    const res = generarBufferPickingConsolidado(seleccionados);
    if (!res.success || !res.buffer || !res.filename) {
      return NextResponse.json(
        {
          error: res.error ?? 'No se pudo generar el picking consolidado.',
          excluidosTalle: res.excluidosTalle,
          excluidosSinLineas: res.excluidosSinLineas
        },
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
        'X-Mibi-Incluidos': String(res.incluidos),
        'X-Mibi-Unidades': String(res.totalUnidades),
        'X-Mibi-Excluidos-Talle': encodeURIComponent(JSON.stringify(res.excluidosTalle)),
        'X-Mibi-Excluidos-Sin-Lineas': encodeURIComponent(JSON.stringify(res.excluidosSinLineas)),
        'Cache-Control': 'no-store'
      }
    });
  } catch (err) {
    if (err instanceof AuthError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json({ error: 'Error al generar Picking Consolidado.' }, { status: 500 });
  }
}
