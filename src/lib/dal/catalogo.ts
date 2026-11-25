import 'server-only';

import { randomUUID } from 'node:crypto';
import { desc, eq } from 'drizzle-orm';
import { CATALOGO_BY_CODIGO, CATALOGO_MIBI } from '@/lib/catalogo-mibi';
import { getDb, isLocalFallbackMode, readLocalStore, writeLocalStore } from '@/lib/db';
import { productosCustom } from '@/lib/db/schema';
import { normalizarTexto } from '@/lib/excel-parser';
import { round2 } from '@/lib/finanzas';
import {
  isDepositoRole,
  type ProductoCatalogo,
  type ProductoCustomRecord,
  type VarianteCatalogo
} from '@/lib/types';
import { AuthError, requireAdmin, requireSession } from './auth';

export interface VarianteInput {
  codigo?: string;
  etiqueta: string;
  aliases?: string[];
}

export interface CrearProductoCustomInput {
  nombre: string;
  rubro: string;
  precioNeto: number;
  variantes: VarianteInput[];
}

export interface EditarProductoCustomInput {
  nombre: string;
  rubro: string;
  precioNeto: number;
  variantes: VarianteInput[];
}

const buildAliasesForVariant = (codigo: string, etiqueta: string, extra: string[] = []): string[] => {
  const tokens = new Set<string>();
  const add = (val: string): void => {
    const norm = normalizarTexto(val);
    if (norm) tokens.add(norm);
  };
  add(codigo);
  add(etiqueta);
  for (const part of etiqueta.split(/[()\/,\-·]+/)) {
    add(part);
  }
  for (const ex of extra) {
    add(ex);
  }
  return Array.from(tokens);
};

const normalizeCode = (raw: string): string =>
  raw
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '-');

const collectUsedCodes = (
  customRecords: ReadonlyArray<ProductoCustomRecord>,
  excludeCustomId?: string
): Set<string> => {
  const used = new Set<string>();
  for (const code of CATALOGO_BY_CODIGO.keys()) {
    used.add(code.toUpperCase());
  }
  for (const rec of customRecords) {
    if (!rec.activo) continue;
    if (excludeCustomId && rec.id === excludeCustomId) continue;
    for (const v of rec.variantes) {
      used.add(v.codigo.toUpperCase());
    }
  }
  return used;
};

const generateNextExtCode = (usedCodes: Set<string>): string => {
  let maxNum = 0;
  for (const code of usedCodes) {
    const m = code.match(/^EXT-(\d+)$/i);
    if (m && m[1]) {
      const n = parseInt(m[1], 10);
      if (Number.isFinite(n) && n > maxNum) maxNum = n;
    }
  }
  let candidateNum = maxNum + 1;
  while (usedCodes.has(`EXT-${String(candidateNum).padStart(3, '0')}`)) {
    candidateNum += 1;
  }
  const nextCode = `EXT-${String(candidateNum).padStart(3, '0')}`;
  usedCodes.add(nextCode);
  return nextCode;
};

const prepareVariantes = (
  rawVariantes: ReadonlyArray<VarianteInput>,
  usedCodes: Set<string>
): VarianteCatalogo[] => {
  if (rawVariantes.length === 0) {
    throw new Error('Definí al menos una variante para el producto.');
  }
  const localCodes = new Set<string>();
  return rawVariantes.map((v, idx) => {
    const etiqueta = v.etiqueta.trim() || (rawVariantes.length === 1 ? 'Único' : `Opción ${idx + 1}`);
    const manualCode = v.codigo ? normalizeCode(v.codigo) : '';
    const codigo = manualCode || generateNextExtCode(usedCodes);

    if (localCodes.has(codigo) || usedCodes.has(codigo)) {
      if (manualCode) {
        throw new Error(`El código "${codigo}" ya existe en el catálogo.`);
      }
    }
    localCodes.add(codigo);
    usedCodes.add(codigo);

    return {
      codigo,
      etiqueta,
      aliases: buildAliasesForVariant(codigo, etiqueta, v.aliases ?? [])
    };
  });
};

const mapRowToCustomRecord = (r: typeof productosCustom.$inferSelect): ProductoCustomRecord => ({
  id: r.id,
  idGrupo: r.idGrupo,
  nombre: r.nombre,
  rubro: r.rubro,
  precioNeto: Number(r.precioNeto) || 0,
  variantes: Array.isArray(r.variantesJson) ? r.variantesJson : [],
  activo: r.activo,
  createdBy: r.createdBy ?? null,
  createdAt: r.createdAt.getTime(),
  updatedAt: r.updatedAt.getTime()
});

export const customRecordToProductoCatalogo = (rec: ProductoCustomRecord): ProductoCatalogo => ({
  idGrupo: rec.idGrupo,
  nombre: rec.nombre,
  rubro: rec.rubro,
  precioNeto: rec.precioNeto,
  variantes: rec.variantes,
  esCustom: true,
  activo: rec.activo,
  customId: rec.id
});

export const listarProductosCustomDal = async (includeArchived = false): Promise<ProductoCustomRecord[]> => {
  await requireSession();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const list = includeArchived
      ? store.productosCustom
      : store.productosCustom.filter((p) => p.activo);
    return [...list].sort((a, b) => b.createdAt - a.createdAt);
  }

  const db = getDb();
  const rows = includeArchived
    ? await db.select().from(productosCustom).orderBy(desc(productosCustom.createdAt))
    : await db
        .select()
        .from(productosCustom)
        .where(eq(productosCustom.activo, true))
        .orderBy(desc(productosCustom.createdAt));

  return rows.map(mapRowToCustomRecord);
};

export const getCatalogoCompletoDal = async (): Promise<ProductoCatalogo[]> => {
  const customActivos = await listarProductosCustomDal(false);
  return [...CATALOGO_MIBI, ...customActivos.map(customRecordToProductoCatalogo)];
};

export const crearProductoCatalogoDal = async (
  input: CrearProductoCustomInput
): Promise<ProductoCustomRecord> => {
  const session = await requireSession();
  if (isDepositoRole(session.rol)) {
    throw new AuthError('Las cuentas de Depósito no pueden crear productos en el catálogo.', 403);
  }

  const nombre = input.nombre.trim();
  const rubro = input.rubro.trim().toUpperCase();
  const precioNeto = round2(Math.max(0, input.precioNeto));
  if (!nombre) throw new Error('Ingresá el nombre del producto.');
  if (!rubro) throw new Error('Ingresá el rubro del producto.');

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const usedCodes = collectUsedCodes(store.productosCustom);
    const variantes = prepareVariantes(input.variantes, usedCodes);
    const idGrupo =
      variantes.length === 1
        ? `CUSTOM-${variantes[0]!.codigo}`
        : `CUSTOM-${variantes[0]!.codigo}-${variantes[variantes.length - 1]!.codigo}`;
    const finalGrupo = store.productosCustom.some((p) => p.idGrupo === idGrupo)
      ? `${idGrupo}-${Date.now().toString(36).toUpperCase()}`
      : idGrupo;

    const now = Date.now();
    const nuevo: ProductoCustomRecord = {
      id: randomUUID(),
      idGrupo: finalGrupo,
      nombre,
      rubro,
      precioNeto,
      variantes,
      activo: true,
      createdBy: session.userId,
      createdAt: now,
      updatedAt: now
    };
    store.productosCustom.unshift(nuevo);
    writeLocalStore(store);
    return nuevo;
  }

  const db = getDb();
  const existingRows = await db.select().from(productosCustom);
  const existingRecords = existingRows.map(mapRowToCustomRecord);
  const usedCodes = collectUsedCodes(existingRecords);
  const variantes = prepareVariantes(input.variantes, usedCodes);

  const baseGrupo =
    variantes.length === 1
      ? `CUSTOM-${variantes[0]!.codigo}`
      : `CUSTOM-${variantes[0]!.codigo}-${variantes[variantes.length - 1]!.codigo}`;
  const finalGrupo = existingRecords.some((p) => p.idGrupo === baseGrupo)
    ? `${baseGrupo}-${Date.now().toString(36).toUpperCase()}`
    : baseGrupo;

  const [created] = await db
    .insert(productosCustom)
    .values({
      idGrupo: finalGrupo,
      nombre,
      rubro,
      precioNeto: precioNeto.toFixed(2),
      variantesJson: variantes,
      activo: true,
      createdBy: session.userId
    })
    .returning();

  if (!created) {
    throw new Error('No se pudo guardar el producto en el catálogo.');
  }
  return mapRowToCustomRecord(created);
};

export const editarProductoCatalogoDal = async (
  id: string,
  input: EditarProductoCustomInput
): Promise<ProductoCustomRecord> => {
  await requireAdmin();

  const nombre = input.nombre.trim();
  const rubro = input.rubro.trim().toUpperCase();
  const precioNeto = round2(Math.max(0, input.precioNeto));
  if (!nombre) throw new Error('Ingresá el nombre del producto.');
  if (!rubro) throw new Error('Ingresá el rubro del producto.');

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const idx = store.productosCustom.findIndex((p) => p.id === id);
    if (idx === -1) throw new Error('Producto no encontrado.');
    const usedCodes = collectUsedCodes(store.productosCustom, id);
    const variantes = prepareVariantes(input.variantes, usedCodes);
    const updated: ProductoCustomRecord = {
      ...store.productosCustom[idx]!,
      nombre,
      rubro,
      precioNeto,
      variantes,
      updatedAt: Date.now()
    };
    store.productosCustom[idx] = updated;
    writeLocalStore(store);
    return updated;
  }

  const db = getDb();
  const existingRows = await db.select().from(productosCustom);
  const existingRecords = existingRows.map(mapRowToCustomRecord);
  const target = existingRecords.find((p) => p.id === id);
  if (!target) throw new Error('Producto no encontrado.');

  const usedCodes = collectUsedCodes(existingRecords, id);
  const variantes = prepareVariantes(input.variantes, usedCodes);

  const [updated] = await db
    .update(productosCustom)
    .set({
      nombre,
      rubro,
      precioNeto: precioNeto.toFixed(2),
      variantesJson: variantes,
      updatedAt: new Date()
    })
    .where(eq(productosCustom.id, id))
    .returning();

  if (!updated) throw new Error('No se pudo actualizar el producto.');
  return mapRowToCustomRecord(updated);
};

export const toggleArchivoProductoCatalogoDal = async (
  id: string,
  activo: boolean
): Promise<ProductoCustomRecord> => {
  await requireAdmin();

  if (isLocalFallbackMode()) {
    const store = readLocalStore();
    const idx = store.productosCustom.findIndex((p) => p.id === id);
    if (idx === -1) throw new Error('Producto no encontrado.');
    const current = store.productosCustom[idx]!;
    if (activo) {
      const usedCodes = collectUsedCodes(store.productosCustom, id);
      for (const v of current.variantes) {
        if (usedCodes.has(v.codigo.toUpperCase())) {
          throw new Error(`No se puede reactivar porque el código "${v.codigo}" está en uso.`);
        }
      }
    }
    const updated: ProductoCustomRecord = {
      ...current,
      activo,
      updatedAt: Date.now()
    };
    store.productosCustom[idx] = updated;
    writeLocalStore(store);
    return updated;
  }

  const db = getDb();
  const existingRows = await db.select().from(productosCustom);
  const existingRecords = existingRows.map(mapRowToCustomRecord);
  const current = existingRecords.find((p) => p.id === id);
  if (!current) throw new Error('Producto no encontrado.');

  if (activo) {
    const usedCodes = collectUsedCodes(existingRecords, id);
    for (const v of current.variantes) {
      if (usedCodes.has(v.codigo.toUpperCase())) {
        throw new Error(`No se puede reactivar porque el código "${v.codigo}" está en uso.`);
      }
    }
  }

  const [updated] = await db
    .update(productosCustom)
    .set({ activo, updatedAt: new Date() })
    .where(eq(productosCustom.id, id))
    .returning();

  if (!updated) throw new Error('No se pudo actualizar el estado del producto.');
  return mapRowToCustomRecord(updated);
};
