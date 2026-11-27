import { CATALOGO_EXTENDIDO_MIBI } from '@/lib/catalogo-mibi';
import { parsearExcelLista006 } from '@/lib/excel-parser';
import type { ProductoCatalogo, ResultadoAnalisisArchivo } from '@/lib/types';

export const analizarPlanilla = (params: {
  buffer: ArrayBuffer;
  nombre: string;
  catalogo?: ReadonlyArray<ProductoCatalogo>;
  hojaSeleccionada?: string;
}): ResultadoAnalisisArchivo => {
  const { buffer, nombre, catalogo = CATALOGO_EXTENDIDO_MIBI, hojaSeleccionada } = params;
  const res = parsearExcelLista006(buffer, nombre, catalogo, { hojaSeleccionada });

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
