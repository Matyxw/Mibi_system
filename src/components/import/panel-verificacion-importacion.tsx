'use client';

import React from 'react';
import {
  AlertTriangle,
  Check,
  CheckCircle2,
  Eye,
  FileText,
  Info,
  RotateCcw,
  Sparkles
} from 'lucide-react';
import { formatARS } from '@/lib/finanzas';
import type {
  CampoAutocompletadoImport,
  ControlImportacionBorrador,
  CorreccionAutomatica,
  TipoArchivoDetectado
} from '@/lib/types';

const LABELS_TIPO: Record<TipoArchivoDetectado, string> = {
  LISTA_006: 'Lista N° 006',
  PLANILLA_PERSONALIZADA: 'Planilla personalizada',
  CSV: 'Planilla CSV',
  FACTURA_AFIP_PDF: 'Factura AFIP',
  PEDIDO_PDF: 'Pedido PDF',
  ALTA_CLIENTE_DOCX: 'Alta de Cliente',
  PEDIDO_DOCX: 'Pedido en Word',
  IMAGEN_O_PDF_ESCANEADO: 'Imagen de referencia',
  PLANILLA_GENERAL_ADMIN: 'Planilla general',
  NO_COMPATIBLE: 'No compatible'
};

const LABELS_CAMPO: Record<CampoAutocompletadoImport['campo'], string> = {
  cliente: 'Cliente',
  whatsapp: 'WhatsApp',
  observaciones: 'Observaciones',
  descuentoPorc: 'Descuento',
  fechaPedido: 'Fecha'
};

interface PanelVerificacionImportacionProps {
  control: ControlImportacionBorrador;
  unidadesActualesBorrador: number;
  alertasActivasCount: number;
  visionDisponible: boolean;
  extrayendoVision: boolean;
  onCambiarHoja: (nombreHoja: string) => void;
  onConfirmarDescuadre: () => void;
  onIrAAlertas: () => void;
  onDeshacerCampo: (campo: CampoAutocompletadoImport['campo']) => void;
  onAplicarCampoIgual: (campo: CampoAutocompletadoImport['campo']) => void;
  onAprobarCorrecciones: () => void;
  onCambiarCorreccion: (corr: CorreccionAutomatica) => void;
  onExtraerConVision?: () => void;
  onConfirmarDatosVision: () => void;
}

export function PanelVerificacionImportacion({
  control,
  unidadesActualesBorrador,
  alertasActivasCount,
  visionDisponible,
  extrayendoVision,
  onCambiarHoja,
  onConfirmarDescuadre,
  onIrAAlertas,
  onDeshacerCampo,
  onAplicarCampoIgual,
  onAprobarCorrecciones,
  onCambiarCorreccion,
  onExtraerConVision,
  onConfirmarDatosVision
}: PanelVerificacionImportacionProps): React.JSX.Element {
  const { auditoria } = control;
  const unidadesArchivoVivas = Math.max(0, unidadesActualesBorrador - control.unidadesBaseBorrador);
  const declaradas = auditoria.unidadesDeclaradasArchivo;
  const difUnidadesViva = declaradas !== null ? unidadesArchivoVivas - declaradas : null;
  const hayDifImporte =
    auditoria.diferenciaImporte !== null && Math.abs(auditoria.diferenciaImporte) > 1;
  const hayDescuadreVivo =
    (difUnidadesViva !== null && difUnidadesViva !== 0) || hayDifImporte;
  const bloqueadoPorDescuadre = hayDescuadreVivo && !control.descuadreConfirmado;

  const chipTipo = LABELS_TIPO[control.tipoArchivo];

  return (
    <div className="rounded-2xl border border-[#EAEAEA] dark:border-[#26262B] bg-white dark:bg-[#141417] p-4 space-y-3 text-xs">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap items-center gap-2 min-w-0">
          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-zinc-100 text-[#111111] dark:bg-[#1B1B1F] dark:text-[#F4F4F5] font-semibold">
            <FileText className="w-3.5 h-3.5 text-zinc-500 dark:text-zinc-400" strokeWidth={1.75} />
            {chipTipo}
          </span>
          <span className="text-zinc-500 dark:text-zinc-400 truncate max-w-[240px]">
            {control.nombreArchivo}
          </span>
        </div>

        {control.hojasDisponibles.length > 1 && (
          <div className="flex flex-wrap items-center gap-1.5">
            {control.hojasDisponibles.map((h) => (
              <button
                key={h.nombre}
                type="button"
                onClick={() => onCambiarHoja(h.nombre)}
                className={`px-2.5 py-1 rounded-full text-[11px] font-medium border transition-colors duration-150 active:scale-[0.98] motion-reduce:transition-none focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-zinc-400 ${
                  h.seleccionada
                    ? 'bg-zinc-900 text-white border-zinc-900 dark:bg-zinc-100 dark:text-zinc-950 dark:border-zinc-100'
                    : 'bg-[#FAFAFA] text-zinc-700 border-zinc-200 hover:bg-zinc-100 dark:bg-[#1B1B1F] dark:text-zinc-300 dark:border-[#2E2E35] dark:hover:bg-[#24242B]'
                }`}
              >
                {h.nombre} ({h.unidades} u.)
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Semáforo de Control de Totales (no aplica en Alta de Cliente ni Imagen vacía) */}
      {control.tipoArchivo !== 'ALTA_CLIENTE_DOCX' &&
        !(control.tipoArchivo === 'IMAGEN_O_PDF_ESCANEADO' && unidadesArchivoVivas === 0 && declaradas === null) && (
          <div
            className={`rounded-xl border px-3.5 py-2.5 flex flex-wrap items-center justify-between gap-2.5 ${
              bloqueadoPorDescuadre
                ? 'bg-red-50/80 border-red-200 text-red-900 dark:bg-red-500/10 dark:border-red-500/30 dark:text-red-200'
                : declaradas !== null || auditoria.totalDeclaradoArchivo !== null
                  ? 'bg-emerald-50/70 border-emerald-200 text-emerald-900 dark:bg-emerald-500/10 dark:border-emerald-500/25 dark:text-emerald-200'
                  : 'bg-[#FAFAFA] border-zinc-200/80 text-zinc-600 dark:bg-[#1B1B1F] dark:border-[#26262B] dark:text-zinc-300'
            }`}
          >
            <div className="flex items-start gap-2 min-w-0">
              {bloqueadoPorDescuadre ? (
                <AlertTriangle className="w-4 h-4 text-red-600 dark:text-red-400 shrink-0 mt-0.5" strokeWidth={1.75} />
              ) : declaradas !== null || auditoria.totalDeclaradoArchivo !== null ? (
                <CheckCircle2 className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0 mt-0.5" strokeWidth={1.75} />
              ) : (
                <Info className="w-4 h-4 text-zinc-400 shrink-0 mt-0.5" strokeWidth={1.75} />
              )}
              <div className="space-y-0.5 tabular-nums">
                {bloqueadoPorDescuadre ? (
                  <div className="font-medium">
                    Diferencia detectada:{' '}
                    {declaradas !== null
                      ? `el archivo declara ${declaradas} u. y se leyeron ${unidadesArchivoVivas} u. (${difUnidadesViva! > 0 ? `+${difUnidadesViva}` : difUnidadesViva})`
                      : `el importe calculado (${formatARS(auditoria.subtotalNetoCalculado)}) difiere del total del archivo (${formatARS(auditoria.totalDeclaradoArchivo ?? 0)})`}
                  </div>
                ) : hayDescuadreVivo && control.descuadreConfirmado ? (
                  <div className="font-medium">
                    Diferencia revisada y confirmada manualmente: {unidadesArchivoVivas} unidades ✓
                  </div>
                ) : declaradas !== null ? (
                  <div className="font-medium">
                    Control exacto: {unidadesArchivoVivas} / {declaradas} unidades ✓
                    {auditoria.totalDeclaradoArchivo !== null && !hayDifImporte && (
                      <span className="ml-2 opacity-80">
                        · {formatARS(auditoria.totalDeclaradoArchivo)} neto
                      </span>
                    )}
                  </div>
                ) : auditoria.totalDeclaradoArchivo !== null ? (
                  <div className="font-medium">
                    Control exacto: {unidadesArchivoVivas} u. · {formatARS(auditoria.totalDeclaradoArchivo)} neto ✓
                  </div>
                ) : (
                  <div>
                    El archivo no informa un total de unidades; se leyeron {unidadesArchivoVivas} u.
                  </div>
                )}
              </div>
            </div>

            {bloqueadoPorDescuadre && (
              <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                {alertasActivasCount > 0 && (
                  <button
                    type="button"
                    onClick={onIrAAlertas}
                    className="px-2.5 py-1 rounded-full text-[11px] font-medium bg-white border border-red-200 text-red-700 hover:bg-red-50 dark:bg-[#1B1B1F] dark:border-red-500/30 dark:text-red-300 transition-colors duration-150 active:scale-[0.98]"
                  >
                    Revisar alertas
                  </button>
                )}
                <button
                  type="button"
                  onClick={onConfirmarDescuadre}
                  className="px-2.5 py-1 rounded-full text-[11px] font-medium bg-red-600 hover:bg-red-700 text-white transition-colors duration-150 active:scale-[0.98]"
                >
                  Confirmar que las cantidades leídas son las correctas
                </button>
              </div>
            )}
          </div>
        )}

      {/* Avisos del análisis */}
      {control.avisos.length > 0 && (
        <div className="rounded-xl bg-zinc-50 dark:bg-[#1B1B1F] border border-zinc-200/80 dark:border-[#26262B] px-3.5 py-2.5 space-y-1 text-zinc-600 dark:text-zinc-300">
          {control.avisos.map((av, i) => (
            <div key={i} className="flex items-start gap-2">
              <Info className="w-3.5 h-3.5 text-zinc-400 shrink-0 mt-0.5" strokeWidth={1.75} />
              <span>{av}</span>
            </div>
          ))}
        </div>
      )}

      {/* Barra de Datos del cliente autocompletados */}
      {control.camposAutocompletados.length > 0 && (
        <div className="rounded-xl bg-[#FAFAFA] dark:bg-[#1B1B1F] border border-zinc-200/80 dark:border-[#26262B] px-3.5 py-2.5 space-y-2">
          <div className="text-[11px] font-medium text-zinc-500 dark:text-zinc-400">
            Datos del cliente completados desde el archivo
          </div>
          <div className="flex flex-wrap gap-1.5">
            {control.camposAutocompletados.map((c) => {
              const habiaValorDistinto = Boolean(c.valorAnterior && c.valorAnterior !== c.valorNuevo);
              return (
                <div
                  key={c.campo}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white dark:bg-[#141417] border border-zinc-200/80 dark:border-[#2E2E35] text-[11px]"
                >
                  <span className="text-zinc-400 dark:text-zinc-500">{LABELS_CAMPO[c.campo]}:</span>
                  <span className="font-medium text-[#111111] dark:text-[#F4F4F5] max-w-[160px] truncate">
                    {c.valorNuevo}
                  </span>
                  {habiaValorDistinto ? (
                    <button
                      type="button"
                      onClick={() => onAplicarCampoIgual(c.campo)}
                      className="text-[10px] font-medium text-zinc-700 hover:text-black dark:text-zinc-300 dark:hover:text-white underline underline-offset-2"
                    >
                      Aplicar igual
                    </button>
                  ) : (
                    <button
                      type="button"
                      onClick={() => onDeshacerCampo(c.campo)}
                      title="Deshacer"
                      aria-label={`Deshacer ${LABELS_CAMPO[c.campo]}`}
                      className="inline-flex items-center gap-0.5 text-[10px] text-zinc-400 hover:text-zinc-800 dark:hover:text-zinc-200"
                    >
                      <RotateCcw className="w-2.5 h-2.5" strokeWidth={1.75} />
                      Deshacer
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Bloque ámbar de Correcciones automáticas transparentes */}
      {control.correcciones.length > 0 && (
        <div
          className={`rounded-xl border px-3.5 py-2.5 space-y-2 ${
            control.correccionesRevisadas
              ? 'bg-zinc-50 border-zinc-200 dark:bg-[#1B1B1F] dark:border-[#26262B] text-zinc-600 dark:text-zinc-300'
              : 'bg-amber-50/70 border-amber-200/80 dark:bg-amber-500/10 dark:border-amber-500/25 text-amber-900 dark:text-amber-200'
          }`}
        >
          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="text-[11px] font-semibold">
              {control.correccionesRevisadas
                ? `Correcciones automáticas revisadas (${control.correcciones.length})`
                : `Correcciones sugeridas aplicadas automáticamente (${control.correcciones.length})`}
            </span>
            {!control.correccionesRevisadas && (
              <button
                type="button"
                onClick={onAprobarCorrecciones}
                className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-medium bg-zinc-900 text-white hover:bg-zinc-800 dark:bg-zinc-100 dark:text-zinc-950 transition-colors duration-150 active:scale-[0.98]"
              >
                <Check className="w-3 h-3" strokeWidth={1.75} />
                Está bien
              </button>
            )}
          </div>
          <div className="space-y-1.5">
            {control.correcciones.map((corr) => (
              <div
                key={corr.id}
                className="flex flex-wrap items-center justify-between gap-2 rounded-lg bg-white/80 dark:bg-[#141417] border border-amber-200/50 dark:border-[#2E2E35] px-2.5 py-1.5 text-[11px]"
              >
                <div className="min-w-0">
                  <span className="font-medium text-[#111111] dark:text-[#F4F4F5]">
                    Fila {corr.fila}: {corr.original} → {corr.corregidoA} {corr.productoNombre}
                  </span>
                  <span className="ml-2 text-zinc-500 dark:text-zinc-400">({corr.motivo})</span>
                </div>
                {!control.correccionesRevisadas && (
                  <button
                    type="button"
                    onClick={() => onCambiarCorreccion(corr)}
                    className="px-2 py-0.5 rounded-full border border-zinc-200 bg-white hover:bg-zinc-50 text-zinc-700 dark:bg-[#1B1B1F] dark:border-[#2E2E35] dark:text-zinc-200 font-medium shrink-0"
                  >
                    Cambiar producto
                  </button>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Opciones para imagen / PDF escaneado */}
      {control.tipoArchivo === 'IMAGEN_O_PDF_ESCANEADO' && (
        <div className="flex flex-wrap items-center justify-between gap-2 pt-0.5">
          {visionDisponible && onExtraerConVision && !control.requiereConfirmacionHumana && (
            <button
              type="button"
              disabled={extrayendoVision}
              onClick={onExtraerConVision}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[11px] font-medium bg-zinc-900 hover:bg-zinc-800 text-white dark:bg-zinc-100 dark:text-zinc-950 disabled:opacity-50 transition-colors duration-150 active:scale-[0.98]"
            >
              <Sparkles className="w-3.5 h-3.5" strokeWidth={1.75} />
              {extrayendoVision ? 'Extrayendo datos...' : 'Extraer datos de la imagen'}
            </button>
          )}
          {control.requiereConfirmacionHumana && !control.confirmacionHumanaOk && (
            <div className="w-full rounded-xl bg-amber-50/70 border border-amber-200/80 dark:bg-amber-500/10 dark:border-amber-500/25 px-3.5 py-2 flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-1.5 text-[11px] text-amber-900 dark:text-amber-200 font-medium">
                <Eye className="w-3.5 h-3.5 shrink-0" strokeWidth={1.75} />
                Compará los datos extraídos contra la imagen de al lado antes de confirmar.
              </div>
              <button
                type="button"
                onClick={onConfirmarDatosVision}
                className="inline-flex items-center gap-1 px-3 py-1 rounded-full text-[11px] font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors duration-150 active:scale-[0.98]"
              >
                <Check className="w-3 h-3" strokeWidth={1.75} />
                Confirmar datos verificados
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
