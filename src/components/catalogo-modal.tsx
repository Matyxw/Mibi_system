'use client';

import React, { useEffect, useState } from 'react';
import { Archive, Edit3, PackagePlus, Plus, RotateCcw, Trash2, X } from 'lucide-react';
import {
  archivarProductoCatalogoAction,
  crearProductoCatalogoAction,
  editarProductoCatalogoAction,
  listarProductosCustomAction
} from '@/app/actions/catalogo-actions';
import { CATALOGO_MIBI } from '@/lib/catalogo-mibi';
import { formatARS, round2 } from '@/lib/finanzas';
import type { ProductoCustomRecord, ToneStatus } from '@/lib/types';
import { StatusPill } from './status-pill';

interface VarianteDraft {
  codigo: string;
  etiqueta: string;
}

interface CatalogoModalProps {
  esAdmin: boolean;
  onClose: () => void;
  onCatalogoActualizado: () => Promise<void>;
  onNotify: (texto: string, tone?: ToneStatus) => void;
}

const RUBROS_BASE = Array.from(new Set([...CATALOGO_MIBI.map((p) => p.rubro), 'PILUSOS', 'ACCESORIOS']));

export function CatalogoModal({
  esAdmin,
  onClose,
  onCatalogoActualizado,
  onNotify
}: CatalogoModalProps): React.JSX.Element {
  const [productos, setProductos] = useState<ProductoCustomRecord[]>([]);
  const [cargando, setCargando] = useState<boolean>(true);
  const [guardando, setGuardando] = useState<boolean>(false);
  const [mostrarForm, setMostrarForm] = useState<boolean>(false);
  const [editId, setEditId] = useState<string | null>(null);

  const [nombre, setNombre] = useState<string>('');
  const [rubroSel, setRubroSel] = useState<string>('PILUSOS');
  const [rubroCustom, setRubroCustom] = useState<string>('');
  const [precioNetoStr, setPrecioNetoStr] = useState<string>('');
  const [tipoVariantes, setTipoVariantes] = useState<'unico' | 'multiple'>('unico');
  const [codigoUnico, setCodigoUnico] = useState<string>('');
  const [variantes, setVariantes] = useState<VarianteDraft[]>([
    { codigo: '', etiqueta: 'T1 (1-3 años)' },
    { codigo: '', etiqueta: 'T2 (3-6 años)' }
  ]);

  const cargar = async (): Promise<void> => {
    setCargando(true);
    const res = await listarProductosCustomAction(esAdmin);
    if (res.ok && res.data) {
      setProductos(res.data);
    }
    setCargando(false);
  };

  useEffect(() => {
    void cargar();
  }, [esAdmin]);

  const resetForm = (): void => {
    setEditId(null);
    setNombre('');
    setRubroSel('PILUSOS');
    setRubroCustom('');
    setPrecioNetoStr('');
    setTipoVariantes('unico');
    setCodigoUnico('');
    setVariantes([
      { codigo: '', etiqueta: 'T1 (1-3 años)' },
      { codigo: '', etiqueta: 'T2 (3-6 años)' }
    ]);
  };

  const abrirEdicion = (p: ProductoCustomRecord): void => {
    setEditId(p.id);
    setNombre(p.nombre);
    if (RUBROS_BASE.includes(p.rubro)) {
      setRubroSel(p.rubro);
      setRubroCustom('');
    } else {
      setRubroSel('__OTRO__');
      setRubroCustom(p.rubro);
    }
    setPrecioNetoStr(String(p.precioNeto));
    if (p.variantes.length <= 1) {
      setTipoVariantes('unico');
      setCodigoUnico(p.variantes[0]?.codigo ?? '');
    } else {
      setTipoVariantes('multiple');
      setVariantes(p.variantes.map((v) => ({ codigo: v.codigo, etiqueta: v.etiqueta })));
    }
    setMostrarForm(true);
  };

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>): Promise<void> => {
    e.preventDefault();
    const rubroFinal = (rubroSel === '__OTRO__' ? rubroCustom : rubroSel).trim().toUpperCase();
    const precio = round2(parseFloat(precioNetoStr.replace(',', '.')) || 0);
    if (!nombre.trim()) {
      onNotify('Ingresá el nombre del producto.', 'amber');
      return;
    }
    if (!rubroFinal) {
      onNotify('Ingresá el rubro.', 'amber');
      return;
    }
    if (precio < 0) {
      onNotify('El precio neto debe ser mayor o igual a 0.', 'amber');
      return;
    }

    const varsPayload =
      tipoVariantes === 'unico'
        ? [{ codigo: codigoUnico.trim() || undefined, etiqueta: 'Único' }]
        : variantes
            .filter((v) => v.etiqueta.trim().length > 0)
            .map((v) => ({
              codigo: v.codigo.trim() || undefined,
              etiqueta: v.etiqueta.trim()
            }));

    if (varsPayload.length === 0) {
      onNotify('Cargá al menos una variante con nombre.', 'amber');
      return;
    }

    setGuardando(true);
    const res = editId
      ? await editarProductoCatalogoAction(editId, {
          nombre: nombre.trim(),
          rubro: rubroFinal,
          precioNeto: precio,
          variantes: varsPayload
        })
      : await crearProductoCatalogoAction({
          nombre: nombre.trim(),
          rubro: rubroFinal,
          precioNeto: precio,
          variantes: varsPayload
        });
    setGuardando(false);

    if (!res.ok) {
      onNotify(res.error ?? 'No se pudo guardar el producto.', 'red');
      return;
    }

    onNotify(editId ? 'Producto actualizado en catálogo.' : 'Producto agregado al catálogo.', 'green');
    resetForm();
    setMostrarForm(false);
    await cargar();
    await onCatalogoActualizado();
  };

  const handleToggleArchivar = async (p: ProductoCustomRecord): Promise<void> => {
    const res = await archivarProductoCatalogoAction({ id: p.id, activo: !p.activo });
    if (!res.ok) {
      onNotify(res.error ?? 'No se pudo actualizar el estado.', 'red');
      return;
    }
    onNotify(p.activo ? `${p.nombre} archivado.` : `${p.nombre} reactivado.`, 'neutral');
    await cargar();
    await onCatalogoActualizado();
  };

  const precioPreview = round2(parseFloat(precioNetoStr.replace(',', '.')) || 0);
  const precioConIva = round2(precioPreview * 1.21);

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/30 px-4"
      onClick={onClose}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        className="w-full max-w-2xl bg-white rounded-2xl border border-[#EAEAEA] shadow-2xl p-6 space-y-5 max-h-[88vh] overflow-y-auto"
      >
        <div className="flex items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <PackagePlus className="w-4 h-4 text-zinc-700" strokeWidth={1.5} />
            <div>
              <h2 className="text-sm font-semibold text-[#111111]">
                Catálogo fuera de planilla ({productos.length})
              </h2>
              <p className="text-[11px] text-zinc-500">
                Pilusos y artículos adicionales combinados con Lista N° 006 ({CATALOGO_MIBI.length} grupos oficiales).
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => {
                if (mostrarForm) {
                  resetForm();
                  setMostrarForm(false);
                } else {
                  resetForm();
                  setMostrarForm(true);
                }
              }}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium bg-zinc-900 hover:bg-zinc-800 text-white"
            >
              <Plus className="w-3.5 h-3.5" strokeWidth={1.75} />
              {mostrarForm ? 'Ver lista' : 'Nuevo producto'}
            </button>
            <button
              type="button"
              onClick={onClose}
              className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-900"
              aria-label="Cerrar catálogo"
            >
              <X className="w-4 h-4" strokeWidth={1.5} />
            </button>
          </div>
        </div>

        {mostrarForm && (
          <form
            onSubmit={(e) => void handleSubmit(e)}
            className="bg-[#FAFAFA] rounded-2xl border border-zinc-200/80 p-4 space-y-3.5 text-xs"
          >
            <div className="font-semibold text-[#111111]">
              {editId ? 'Editar producto fuera de planilla' : 'Agregar producto fuera de planilla'}
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="block text-[11px] text-zinc-500 mb-1">Nombre del producto *</label>
                <input
                  type="text"
                  value={nombre}
                  onChange={(e) => setNombre(e.target.value)}
                  placeholder="Ej. Piluso Shark"
                  className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                />
              </div>

              <div>
                <label className="block text-[11px] text-zinc-500 mb-1">Rubro *</label>
                <div className="flex gap-2">
                  <select
                    value={rubroSel}
                    onChange={(e) => setRubroSel(e.target.value)}
                    className="flex-1 px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                  >
                    {RUBROS_BASE.map((r) => (
                      <option key={r} value={r}>
                        {r}
                      </option>
                    ))}
                    <option value="__OTRO__">Otro rubro...</option>
                  </select>
                  {rubroSel === '__OTRO__' && (
                    <input
                      type="text"
                      value={rubroCustom}
                      onChange={(e) => setRubroCustom(e.target.value)}
                      placeholder="Ej. PILUSOS"
                      className="w-32 px-3 py-2 bg-white border border-zinc-200 rounded-xl uppercase"
                    />
                  )}
                </div>
              </div>

              <div>
                <label className="block text-[11px] text-zinc-500 mb-1">Precio neto unitario ($) *</label>
                <input
                  type="number"
                  min={0}
                  step="0.01"
                  value={precioNetoStr}
                  onChange={(e) => setPrecioNetoStr(e.target.value)}
                  placeholder="0.00"
                  className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl tabular-nums"
                />
                <div className="text-[11px] text-zinc-400 mt-1 tabular-nums">
                  Referencia con IVA 21%: {formatARS(precioConIva)}
                </div>
              </div>

              <div>
                <label className="block text-[11px] text-zinc-500 mb-1">Variantes</label>
                <div className="grid grid-cols-2 gap-1.5">
                  <button
                    type="button"
                    onClick={() => setTipoVariantes('unico')}
                    className={`py-2 px-3 rounded-xl border font-medium ${
                      tipoVariantes === 'unico'
                        ? 'bg-zinc-900 text-white border-zinc-900'
                        : 'bg-white text-zinc-700 border-zinc-200'
                    }`}
                  >
                    Único
                  </button>
                  <button
                    type="button"
                    onClick={() => setTipoVariantes('multiple')}
                    className={`py-2 px-3 rounded-xl border font-medium ${
                      tipoVariantes === 'multiple'
                        ? 'bg-zinc-900 text-white border-zinc-900'
                        : 'bg-white text-zinc-700 border-zinc-200'
                    }`}
                  >
                    Con talles o colores
                  </button>
                </div>
              </div>
            </div>

            {tipoVariantes === 'unico' ? (
              <div>
                <label className="block text-[11px] text-zinc-500 mb-1">
                  Código (opcional — si queda vacío se genera automático)
                </label>
                <input
                  type="text"
                  value={codigoUnico}
                  onChange={(e) => setCodigoUnico(e.target.value)}
                  placeholder="Ej. PIL-01 (o vacío → EXT-001)"
                  className="w-full px-3 py-2 bg-white border border-zinc-200 rounded-xl tabular-nums uppercase"
                />
              </div>
            ) : (
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-zinc-500">Talles / colores (2 o más)</span>
                  <button
                    type="button"
                    onClick={() => setVariantes((prev) => [...prev, { codigo: '', etiqueta: '' }])}
                    className="text-[11px] font-medium text-zinc-800 hover:underline inline-flex items-center gap-1"
                  >
                    <Plus className="w-3 h-3" strokeWidth={1.75} /> Agregar opción
                  </button>
                </div>
                {variantes.map((v, idx) => (
                  <div key={idx} className="flex items-center gap-2">
                    <input
                      type="text"
                      value={v.codigo}
                      onChange={(e) =>
                        setVariantes((prev) =>
                          prev.map((item, i) => (i === idx ? { ...item, codigo: e.target.value } : item))
                        )
                      }
                      placeholder="Código (opc.)"
                      className="w-28 px-2.5 py-1.5 bg-white border border-zinc-200 rounded-lg tabular-nums uppercase"
                    />
                    <input
                      type="text"
                      value={v.etiqueta}
                      onChange={(e) =>
                        setVariantes((prev) =>
                          prev.map((item, i) => (i === idx ? { ...item, etiqueta: e.target.value } : item))
                        )
                      }
                      placeholder={`Talle o color #${idx + 1}`}
                      className="flex-1 px-2.5 py-1.5 bg-white border border-zinc-200 rounded-lg"
                    />
                    {variantes.length > 2 && (
                      <button
                        type="button"
                        onClick={() => setVariantes((prev) => prev.filter((_, i) => i !== idx))}
                        className="p-1.5 text-zinc-400 hover:text-red-600"
                        aria-label="Quitar opción"
                      >
                        <Trash2 className="w-3.5 h-3.5" strokeWidth={1.5} />
                      </button>
                    )}
                  </div>
                ))}
              </div>
            )}

            <div className="flex justify-end gap-2 pt-1">
              <button
                type="button"
                onClick={() => {
                  resetForm();
                  setMostrarForm(false);
                }}
                className="px-3.5 py-2 rounded-full border border-zinc-200 bg-white text-zinc-600"
              >
                Cancelar
              </button>
              <button
                type="submit"
                disabled={guardando}
                className="px-4 py-2 rounded-full bg-zinc-900 text-white font-medium hover:bg-zinc-800 disabled:opacity-50"
              >
                {guardando ? 'Guardando...' : editId ? 'Guardar cambios' : 'Guardar producto'}
              </button>
            </div>
          </form>
        )}

        {cargando ? (
          <div className="text-xs text-zinc-400 py-6 text-center">Cargando productos...</div>
        ) : productos.length === 0 ? (
          <div className="text-xs text-zinc-400 py-8 text-center border border-dashed border-zinc-200 rounded-xl">
            Todavía no hay productos agregados fuera de la Lista N° 006.
          </div>
        ) : (
          <div className="divide-y divide-zinc-200/60 border border-zinc-200/80 rounded-xl">
            {productos.map((p) => (
              <div key={p.id} className="px-4 py-3 flex items-center justify-between gap-3 text-xs">
                <div className="min-w-0">
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-[#111111] truncate">{p.nombre}</span>
                    <StatusPill tone={p.activo ? 'green' : 'neutral'}>
                      {p.activo ? p.rubro : 'Archivado'}
                    </StatusPill>
                  </div>
                  <div className="text-[11px] text-zinc-500 mt-0.5 tabular-nums">
                    {p.variantes.map((v) => `${v.codigo} (${v.etiqueta})`).join(' · ')}
                  </div>
                </div>

                <div className="flex items-center gap-2 shrink-0">
                  <span className="font-semibold text-[#111111] tabular-nums">
                    {formatARS(p.precioNeto)} neto
                  </span>
                  {esAdmin && (
                    <>
                      <button
                        type="button"
                        onClick={() => abrirEdicion(p)}
                        className="px-2.5 py-1 rounded-full border border-zinc-200 text-[11px] text-zinc-700 hover:bg-zinc-50 inline-flex items-center gap-1"
                      >
                        <Edit3 className="w-3 h-3" strokeWidth={1.5} />
                        Editar
                      </button>
                      <button
                        type="button"
                        onClick={() => void handleToggleArchivar(p)}
                        className="px-2.5 py-1 rounded-full border border-zinc-200 text-[11px] text-zinc-600 hover:bg-zinc-50 inline-flex items-center gap-1"
                      >
                        {p.activo ? (
                          <>
                            <Archive className="w-3 h-3" strokeWidth={1.5} /> Archivar
                          </>
                        ) : (
                          <>
                            <RotateCcw className="w-3 h-3" strokeWidth={1.5} /> Activar
                          </>
                        )}
                      </button>
                    </>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
