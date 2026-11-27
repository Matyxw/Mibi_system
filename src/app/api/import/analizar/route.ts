import { NextResponse, type NextRequest } from 'next/server';
import { AuthError, requireSession } from '@/lib/dal/auth';
import { getCatalogoCompletoDal } from '@/lib/dal/catalogo';
import { CATALOGO_EXTENDIDO_BY_CODIGO, CATALOGO_EXTENDIDO_MIBI } from '@/lib/catalogo-mibi';
import { analizarArchivo } from '@/lib/import/analizar-archivo';
import { detectarFormatoPorFirma } from '@/lib/import/firma-archivo';
import { extraerConVision } from '@/lib/import/vision-gemini';
import { isDepositoRole, type ProductoCatalogo } from '@/lib/types';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

const MAX_FILE_BYTES = 10 * 1024 * 1024;
const MAX_REQUEST_BYTES = MAX_FILE_BYTES + 64 * 1024;

export async function POST(req: NextRequest): Promise<NextResponse> {
  try {
    const session = await requireSession();
    if (isDepositoRole(session.rol)) {
      return NextResponse.json(
        { error: 'Tu cuenta de Depósito no puede subir archivos de pedidos.' },
        { status: 403 }
      );
    }

    const contentLength = Number(req.headers.get('content-length') ?? '0');
    if (Number.isFinite(contentLength) && contentLength > MAX_REQUEST_BYTES) {
      return NextResponse.json(
        { error: 'El archivo supera el límite máximo permitido de 10 MB.' },
        { status: 413 }
      );
    }

    const formData = await req.formData().catch(() => null);
    if (!formData) {
      return NextResponse.json({ error: 'Solicitud inválida.' }, { status: 400 });
    }

    const archivo = formData.get('archivo');
    if (!(archivo instanceof File)) {
      return NextResponse.json({ error: 'Falta adjuntar el archivo a analizar.' }, { status: 400 });
    }

    if (archivo.size === 0) {
      return NextResponse.json({ error: 'El archivo está vacío.' }, { status: 400 });
    }
    if (archivo.size > MAX_FILE_BYTES) {
      return NextResponse.json(
        { error: 'El archivo supera el límite máximo permitido de 10 MB.' },
        { status: 413 }
      );
    }

    const hojaRaw = formData.get('hoja');
    const hojaSeleccionada = typeof hojaRaw === 'string' && hojaRaw.trim() ? hojaRaw.trim() : undefined;
    const modoRaw = formData.get('modo');
    const modo = modoRaw === 'vision' ? 'vision' : 'estandar';

    const bytes = new Uint8Array(await archivo.arrayBuffer());
    const formato = await detectarFormatoPorFirma(bytes, archivo.name);
    if (formato === 'DESCONOCIDO') {
      return NextResponse.json(
        { error: 'Formato no reconocido. Subí Excel, CSV, PDF, Word o una imagen.' },
        { status: 415 }
      );
    }

    const catalogo = await getCatalogoCompletoDal();

    if (modo === 'vision') {
      if (formato !== 'JPEG' && formato !== 'PNG' && formato !== 'WEBP') {
        return NextResponse.json(
          { error: 'La lectura de imagen solo admite archivos JPG, PNG o WEBP.' },
          { status: 415 }
        );
      }

      const mimeType =
        formato === 'JPEG' ? 'image/jpeg' : formato === 'PNG' ? 'image/png' : 'image/webp';

      const catalogoPorCodigo = new Map<string, ProductoCatalogo>(CATALOGO_EXTENDIDO_BY_CODIGO);
      const mapUnificado = new Map<string, ProductoCatalogo>();
      for (const p of CATALOGO_EXTENDIDO_MIBI) mapUnificado.set(p.idGrupo, p);
      for (const p of catalogo) {
        mapUnificado.set(p.idGrupo, p);
        for (const v of p.variantes) catalogoPorCodigo.set(v.codigo.trim().toUpperCase(), p);
      }

      try {
        const resVision = await extraerConVision({
          bytes,
          mimeType,
          fileName: archivo.name,
          catalogoPorCodigo,
          catalogoUnificado: Array.from(mapUnificado.values())
        });
        return NextResponse.json(resVision, {
          status: 200,
          headers: { 'Cache-Control': 'no-store' }
        });
      } catch (err) {
        const msg = err instanceof Error ? err.message : 'No se pudo analizar la imagen.';
        return NextResponse.json({ error: msg }, { status: 400 });
      }
    }

    const res = await analizarArchivo({
      bytes,
      nombre: archivo.name,
      catalogo,
      hojaSeleccionada
    });

    return NextResponse.json(res, {
      status: 200,
      headers: { 'Cache-Control': 'no-store' }
    });
  } catch (err) {
    if (err instanceof AuthError) {
      return NextResponse.json({ error: err.message }, { status: err.status });
    }
    return NextResponse.json(
      { error: 'No se pudo analizar el archivo en este momento.' },
      { status: 500 }
    );
  }
}
