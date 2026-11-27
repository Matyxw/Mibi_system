'use client';

import React, { useEffect, useState } from 'react';
import { RotateCw, X, ZoomIn, ZoomOut } from 'lucide-react';

interface VisorReferenciaProps {
  archivo: File;
  onCerrar: () => void;
}

export function VisorReferencia({ archivo, onCerrar }: VisorReferenciaProps): React.JSX.Element {
  const [objectUrl, setObjectUrl] = useState<string>('');
  const [zoom, setZoom] = useState<number>(1);
  const [rotacion, setRotacion] = useState<number>(0);

  useEffect(() => {
    const url = URL.createObjectURL(archivo);
    setObjectUrl(url);
    setZoom(1);
    setRotacion(0);
    return () => {
      URL.revokeObjectURL(url);
    };
  }, [archivo]);

  const esPdf =
    archivo.type === 'application/pdf' || archivo.name.toLowerCase().endsWith('.pdf');

  const cycleZoom = (dir: 1 | -1): void => {
    const levels = [1, 1.5, 2] as const;
    const curIdx = levels.indexOf(zoom as 1 | 1.5 | 2);
    if (dir === 1) {
      const next = levels[Math.min(levels.length - 1, (curIdx === -1 ? 0 : curIdx) + 1)] ?? 2;
      setZoom(next);
    } else {
      const next = levels[Math.max(0, (curIdx === -1 ? 0 : curIdx) - 1)] ?? 1;
      setZoom(next);
    }
  };

  return (
    <aside className="rounded-2xl border border-[#EAEAEA] dark:border-[#26262B] bg-white dark:bg-[#141417] shadow-[0_1px_2px_rgba(0,0,0,0.03)] overflow-hidden lg:sticky lg:top-6 flex flex-col max-h-[82vh]">
      <div className="px-4 py-3 border-b border-[#EAEAEA] dark:border-[#26262B] flex items-center justify-between gap-2 bg-[#FAFAFA] dark:bg-[#1B1B1F]">
        <div className="min-w-0">
          <div className="text-xs font-semibold text-[#111111] dark:text-[#F4F4F5]">
            Referencia visual
          </div>
          <div className="text-[11px] text-zinc-500 dark:text-zinc-400 truncate">
            {archivo.name}
          </div>
        </div>

        <div className="flex items-center gap-1 shrink-0">
          {!esPdf && (
            <>
              <button
                type="button"
                onClick={() => cycleZoom(-1)}
                disabled={zoom <= 1}
                aria-label="Alejar"
                className="p-1.5 rounded-lg text-zinc-600 hover:bg-zinc-200/60 dark:text-zinc-300 dark:hover:bg-[#2A2A32] disabled:opacity-35 transition-colors duration-150"
              >
                <ZoomOut className="w-3.5 h-3.5" strokeWidth={1.75} />
              </button>
              <span className="text-[11px] tabular-nums text-zinc-600 dark:text-zinc-300 w-9 text-center">
                {zoom}×
              </span>
              <button
                type="button"
                onClick={() => cycleZoom(1)}
                disabled={zoom >= 2}
                aria-label="Acercar"
                className="p-1.5 rounded-lg text-zinc-600 hover:bg-zinc-200/60 dark:text-zinc-300 dark:hover:bg-[#2A2A32] disabled:opacity-35 transition-colors duration-150"
              >
                <ZoomIn className="w-3.5 h-3.5" strokeWidth={1.75} />
              </button>
              <button
                type="button"
                onClick={() => setRotacion((r) => (r + 90) % 360)}
                aria-label="Rotar 90 grados"
                className="p-1.5 rounded-lg text-zinc-600 hover:bg-zinc-200/60 dark:text-zinc-300 dark:hover:bg-[#2A2A32] transition-colors duration-150"
              >
                <RotateCw className="w-3.5 h-3.5" strokeWidth={1.75} />
              </button>
            </>
          )}
          <button
            type="button"
            onClick={onCerrar}
            aria-label="Cerrar referencia visual"
            className="p-1.5 rounded-lg text-zinc-500 hover:text-zinc-900 hover:bg-zinc-200/60 dark:text-zinc-400 dark:hover:text-white dark:hover:bg-[#2A2A32] transition-colors duration-150"
          >
            <X className="w-3.5 h-3.5" strokeWidth={1.75} />
          </button>
        </div>
      </div>

      <div className="flex-1 overflow-auto p-3 bg-zinc-100/70 dark:bg-[#0D0D10] flex items-center justify-center min-h-[340px]">
        {!objectUrl ? (
          <div className="text-xs text-zinc-400">Cargando vista previa...</div>
        ) : esPdf ? (
          <object
            data={objectUrl}
            type="application/pdf"
            className="w-full h-[70vh] rounded-lg border border-zinc-200 dark:border-[#26262B]"
          >
            <p className="text-xs text-zinc-500 p-4 text-center">
              Tu navegador no pudo incrustar el PDF. Podés abrirlo localmente para consultarlo.
            </p>
          </object>
        ) : (
          <div className="overflow-auto w-full h-full flex items-center justify-center">
            <img
              src={objectUrl}
              alt={archivo.name}
              style={{
                transform: `scale(${zoom}) rotate(${rotacion}deg)`,
                transformOrigin: 'center center'
              }}
              className="max-w-full max-h-[68vh] object-contain rounded-lg transition-transform duration-150 motion-reduce:transition-none"
            />
          </div>
        )}
      </div>
    </aside>
  );
}
