import { CATALOGO_EXTENDIDO_MIBI } from '@/lib/catalogo-mibi';
import { parsearExcelLista006 } from '@/lib/excel-parser';
import { extraerEstilosYComentariosXlsx } from '@/lib/import/xlsx-estilos';
import type { ProductoCatalogo, ResultadoAnalisisArchivo } from '@/lib/types';

export const analizarPlanilla = (params: {
  buffer: ArrayBuffer;
  nombre: string;
  catalogo?: ReadonlyArray<ProductoCatalogo>;
  hojaSeleccionada?: string;
  celdasTachadasPorHoja?: ReadonlyMap<string, ReadonlySet<string>>;
  comentariosThreadedPorHoja?: ReadonlyMap<string, ReadonlyMap<string, string>>;
}): ResultadoAnalisisArchivo => {
  const {
    buffer,
    nombre,
    catalogo = CATALOGO_EXTENDIDO_MIBI,
    hojaSeleccionada,
    celdasTachadasPorHoja,
    comentariosThreadedPorHoja
  } = params;
  const res = parsearExcelLista006(buffer, nombre, catalogo, {
    hojaSeleccionada,
    celdasTachadasPorHoja,
    comentariosThreadedPorHoja
  });

  const avisos: string[] = [];
  if (res.ok && res.hojasDisponibles.length > 1) {
    const sel = res.hojasDisponibles.find((h) => h.seleccionada);
    if (sel) {
      avisos.push(
        `El archivo tiene ${res.hojasDisponibles.length} hojas con artículos. Se está mostrando "${sel.nombre}".`
      );
    }
  }

  return {
    ...res,
    nombreArchivo: nombre,
    avisos,
    visionDisponible: false,
    requiereConfirmacionHumana: false
  };
};

export const analizarPlanillaAsync = async (params: {
  buffer: ArrayBuffer;
  nombre: string;
  catalogo?: ReadonlyArray<ProductoCatalogo>;
  hojaSeleccionada?: string;
}): Promise<ResultadoAnalisisArchivo> => {
  const isCsv = /\.csv$/i.test(params.nombre);
  if (isCsv) {
    return analizarPlanilla(params);
  }
  const estilos = await extraerEstilosYComentariosXlsx(params.buffer);
  return analizarPlanilla({
    ...params,
    celdasTachadasPorHoja: estilos.celdasTachadasPorHoja,
    comentariosThreadedPorHoja: estilos.comentariosThreadedPorHoja
  });
};
