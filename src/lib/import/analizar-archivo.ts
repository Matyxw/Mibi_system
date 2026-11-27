import {
  CATALOGO_EXTENDIDO_BY_CODIGO,
  CATALOGO_EXTENDIDO_MIBI
} from '@/lib/catalogo-mibi';
import { auditoriaVacia } from '@/lib/excel-parser';
import type { ProductoCatalogo, ResultadoAnalisisArchivo } from '@/lib/types';
import { analizarPlanilla } from './analizar-planilla';
import { analizarDocx } from './docx-parser';
import { detectarFormatoPorFirma, extensionCoincide } from './firma-archivo';
import { analizarPdf } from './pdf-parser';

const construirContextoCatalogo = (
  nombre: string,
  catalogo: ReadonlyArray<ProductoCatalogo>
): {
  fileName: string;
  catalogoPorCodigo: Map<string, ProductoCatalogo>;
  catalogoUnificado: ProductoCatalogo[];
} => {
  const catalogoPorCodigo = new Map<string, ProductoCatalogo>(CATALOGO_EXTENDIDO_BY_CODIGO);
  const map = new Map<string, ProductoCatalogo>();
  for (const p of CATALOGO_EXTENDIDO_MIBI) map.set(p.idGrupo, p);
  for (const p of catalogo) {
    map.set(p.idGrupo, p);
    for (const v of p.variantes) {
      catalogoPorCodigo.set(v.codigo.trim().toUpperCase(), p);
    }
  }
  return {
    fileName: nombre,
    catalogoPorCodigo,
    catalogoUnificado: Array.from(map.values())
  };
};

export const analizarArchivo = async (params: {
  bytes: Uint8Array;
  nombre: string;
  catalogo?: ReadonlyArray<ProductoCatalogo>;
  hojaSeleccionada?: string;
}): Promise<ResultadoAnalisisArchivo> => {
  const { bytes, nombre, catalogo = CATALOGO_EXTENDIDO_MIBI, hojaSeleccionada } = params;
  const formato = await detectarFormatoPorFirma(bytes, nombre);
  const checkExt = extensionCoincide(nombre, formato);
  const avisosBase: string[] = !checkExt.coincide && checkExt.aviso ? [checkExt.aviso] : [];

  if (formato === 'DESCONOCIDO') {
    return {
      ok: false,
      errorFatal: 'Formato no reconocido. Subí Excel, CSV, PDF, Word o una imagen.',
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'NO_COMPATIBLE',
      auditoria: auditoriaVacia('BLOQUEADO'),
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados: null,
      nombreArchivo: nombre,
      avisos: avisosBase,
      visionDisponible: false,
      requiereConfirmacionHumana: false
    };
  }

  if (formato === 'JPEG' || formato === 'PNG' || formato === 'WEBP') {
    const aud = auditoriaVacia('REVISAR');
    aud.motivos.push('Imagen cargada como referencia visual.');
    return {
      ok: true,
      lineas: [],
      advertencias: [],
      avisosPrecio: [],
      tipoArchivo: 'IMAGEN_O_PDF_ESCANEADO',
      auditoria: aud,
      correccionesAutomaticas: [],
      hojasDisponibles: [],
      datosDetectados: null,
      nombreArchivo: nombre,
      avisos: [
        ...avisosBase,
        'Guardamos la imagen como referencia: podés verla al lado mientras cargás el pedido.'
      ],
      visionDisponible: Boolean(process.env.GEMINI_API_KEY?.trim()),
      requiereConfirmacionHumana: false
    };
  }

  if (formato === 'XLSX' || formato === 'XLS' || formato === 'CSV') {
    const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const effectiveName = formato === 'CSV' && !/\.csv$/i.test(nombre) ? `${nombre}.csv` : nombre;
    const res = analizarPlanilla({
      buffer: ab,
      nombre: effectiveName,
      catalogo,
      hojaSeleccionada
    });
    return {
      ...res,
      nombreArchivo: nombre,
      avisos: [...avisosBase, ...res.avisos]
    };
  }

  const ctx = construirContextoCatalogo(nombre, catalogo);

  if (formato === 'DOCX') {
    const res = await analizarDocx(bytes, ctx);
    return {
      ...res,
      avisos: [...avisosBase, ...res.avisos]
    };
  }

  // formato === 'PDF'
  const res = await analizarPdf(bytes, ctx);
  return {
    ...res,
    avisos: [...avisosBase, ...res.avisos]
  };
};
