import 'server-only';

import { z } from 'zod';
import {
  analizarFilasTabulares,
  construirAuditoriaDesdeTabular
} from '@/lib/excel-parser';
import type {
  CellPrimitive,
  DatosClienteDetectados,
  ProductoCatalogo,
  ResultadoAnalisisArchivo
} from '@/lib/types';
import { normalizarCuit, validarCuit } from './numeros';

const VisionResponseSchema = z.object({
  tipo: z.enum(['FICHA_CLIENTE', 'COMPROBANTE_PAGO', 'PEDIDO_ARTICULOS', 'OTRO']),
  cliente: z
    .object({
      nombre: z.string().optional(),
      cuit: z.string().optional(),
      whatsapp: z.string().optional(),
      email: z.string().optional(),
      domicilio: z.string().optional(),
      transporte: z.string().optional()
    })
    .optional(),
  comprobante: z
    .object({
      monto: z.number().optional(),
      fecha: z.string().optional(),
      banco: z.string().optional()
    })
    .optional(),
  articulos: z
    .array(
      z.object({
        codigo: z.string().optional(),
        descripcion: z.string().optional(),
        cantidad: z.number().optional(),
        detalle: z.string().optional()
      })
    )
    .default([])
});

export const extraerConVision = async (params: {
  bytes: Uint8Array;
  mimeType: string;
  fileName: string;
  catalogoPorCodigo: ReadonlyMap<string, ProductoCatalogo>;
  catalogoUnificado: ReadonlyArray<ProductoCatalogo>;
}): Promise<ResultadoAnalisisArchivo> => {
  const apiKey = process.env.GEMINI_API_KEY?.trim();
  if (!apiKey) {
    throw new Error('La extracción asistida de imágenes no está habilitada en el servidor.');
  }

  const model = process.env.GEMINI_VISION_MODEL?.trim() || 'gemini-2.5-flash';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;

  const prompt = [
    'Analizá esta imagen comercial de un pedido mayorista argentino.',
    'Clasificá tipo como: FICHA_CLIENTE, COMPROBANTE_PAGO, PEDIDO_ARTICULOS u OTRO.',
    'Devolvé únicamente JSON válido con esta estructura exacta:',
    '{ "tipo": "...", "cliente": { "nombre": "...", "cuit": "...", "whatsapp": "...", "email": "...", "domicilio": "...", "transporte": "..." }, "comprobante": { "monto": 0, "fecha": "YYYY-MM-DD", "banco": "..." }, "articulos": [ { "codigo": "6504", "descripcion": "...", "cantidad": 2, "detalle": "T1 / Rosa" } ] }',
    'No inventes códigos ni importes que no figuren en la imagen.'
  ].join('\n');

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 25_000);

  let rawJsonText = '';
  try {
    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': apiKey
      },
      body: JSON.stringify({
        contents: [
          {
            role: 'user',
            parts: [
              { text: prompt },
              {
                inline_data: {
                  mime_type: params.mimeType,
                  data: Buffer.from(params.bytes).toString('base64')
                }
              }
            ]
          }
        ],
        generationConfig: {
          responseMimeType: 'application/json',
          temperature: 0.1
        }
      }),
      signal: controller.signal
    });

    if (!resp.ok) {
      throw new Error('No se pudo analizar la imagen en este momento.');
    }

    const payload = (await resp.json()) as {
      candidates?: Array<{ content?: { parts?: Array<{ text?: string }> } }>;
    };
    rawJsonText = payload.candidates?.[0]?.content?.parts?.[0]?.text ?? '';
  } finally {
    clearTimeout(timeout);
  }

  let parsedRaw: unknown;
  try {
    parsedRaw = JSON.parse(rawJsonText);
  } catch {
    throw new Error('La respuesta del análisis de imagen no tuvo un formato válido.');
  }

  const valid = VisionResponseSchema.safeParse(parsedRaw);
  if (!valid.success) {
    throw new Error('Los datos extraídos de la imagen no son consistentes.');
  }

  const data = valid.data;
  const rows: CellPrimitive[][] = [['Codigos', 'Producto', 'Cantidades Pedidas', 'Detalle Talles/ Colores']];
  for (const art of data.articulos) {
    if (!art.cantidad || art.cantidad <= 0) continue;
    rows.push([
      art.codigo ?? '',
      art.descripcion ?? '',
      Math.round(art.cantidad),
      art.detalle ?? ''
    ]);
  }

  const tab = analizarFilasTabulares(rows, 'Imagen', {
    catalogoPorCodigo: params.catalogoPorCodigo,
    catalogoUnificado: params.catalogoUnificado,
    fileName: params.fileName
  });

  const auditoria = construirAuditoriaDesdeTabular(tab, []);
  if (auditoria.estadoIntegridad === 'EXACTO') {
    auditoria.estadoIntegridad = 'REVISAR';
    auditoria.motivos.push('Los datos extraídos de la imagen requieren confirmación manual antes de guardar.');
  }

  const cli = data.cliente;
  const cuitOk = cli?.cuit && validarCuit(cli.cuit) ? normalizarCuit(cli.cuit) : undefined;
  const compObs = data.comprobante?.monto
    ? `Comprobante leído: $${data.comprobante.monto}${data.comprobante.fecha ? ` (${data.comprobante.fecha})` : ''}${data.comprobante.banco ? ` · ${data.comprobante.banco}` : ''}`
    : undefined;

  const datosDetectados: DatosClienteDetectados | null =
    cli?.nombre || cuitOk || cli?.whatsapp || compObs
      ? {
          ...(cli?.nombre ? { cliente: cli.nombre } : {}),
          ...(cuitOk ? { cuit: cuitOk } : {}),
          ...(cli?.whatsapp ? { whatsapp: cli.whatsapp } : {}),
          ...(cli?.email ? { email: cli.email } : {}),
          ...(cli?.domicilio ? { domicilio: cli.domicilio } : {}),
          ...(cli?.transporte ? { transporte: cli.transporte } : {}),
          ...(compObs ? { observacionesSugeridas: compObs } : {}),
          origen: 'IMAGEN'
        }
      : null;

  return {
    ok: true,
    lineas: tab.lineas,
    advertencias: tab.advertencias,
    avisosPrecio: tab.avisosPrecio,
    tipoArchivo: 'IMAGEN_O_PDF_ESCANEADO',
    auditoria,
    correccionesAutomaticas: tab.correcciones,
    hojasDisponibles: [],
    datosDetectados,
    comentariosDetectados: [],
    nombreArchivo: params.fileName,
    avisos: ['Revisá los datos extraídos frente a la imagen antes de confirmar.'],
    visionDisponible: true,
    requiereConfirmacionHumana: true,
    tipoImagen: data.tipo
  };
};
