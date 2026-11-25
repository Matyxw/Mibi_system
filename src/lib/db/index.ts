import 'server-only';

import fs from 'node:fs';
import path from 'node:path';
import postgres, { type Sql } from 'postgres';
import { drizzle, type PostgresJsDatabase } from 'drizzle-orm/postgres-js';
import * as schema from './schema';
import type {
  Pedido,
  PedidoAsignacion,
  ProductoCustomRecord,
  RolUsuario
} from '@/lib/types';
import { hashPassword } from '@/lib/auth/password';

export interface LocalUsuarioRecord {
  id: string;
  username: string;
  nombre: string;
  rol: RolUsuario;
  passwordHash: string;
  activo: boolean;
  tokenVersion: number;
  createdAt: number;
  updatedAt: number;
}

export interface LocalDevStore {
  nextNumeroSeq: number;
  nextCustomCodeSeq: number;
  usuarios: LocalUsuarioRecord[];
  pedidos: Pedido[];
  productosCustom: ProductoCustomRecord[];
  pedidoAsignaciones: PedidoAsignacion[];
}

export type DrizzleDb = PostgresJsDatabase<typeof schema>;

const globalForDb = globalThis as unknown as {
  __mibiSql?: Sql;
  __mibiDb?: DrizzleDb;
};

const assertProdEnv = (): void => {
  if (process.env.NODE_ENV === 'production') {
    if (!process.env.DATABASE_URL) {
      throw new Error('FATAL: Falta DATABASE_URL en entorno de producción.');
    }
    if (!process.env.AUTH_SECRET || process.env.AUTH_SECRET.length < 32) {
      throw new Error('FATAL: AUTH_SECRET debe existir y tener al menos 32 bytes en producción.');
    }
  }
};

export const isLocalFallbackMode = (): boolean => {
  assertProdEnv();
  return !process.env.DATABASE_URL && process.env.NODE_ENV !== 'production';
};

export const getDb = (): DrizzleDb => {
  assertProdEnv();
  const databaseUrl = process.env.DATABASE_URL;
  if (!databaseUrl) {
    throw new Error('DATABASE_URL no está configurada.');
  }

  if (!globalForDb.__mibiDb) {
    const sqlClient = postgres(databaseUrl, {
      prepare: false,
      max: 5,
      idle_timeout: 20,
      connect_timeout: 10,
      ssl: 'require'
    });
    globalForDb.__mibiSql = sqlClient;
    globalForDb.__mibiDb = drizzle(sqlClient, { schema });
  }
  return globalForDb.__mibiDb;
};

const LOCAL_DATA_DIR = path.join(process.cwd(), '.local-data');
const LOCAL_DATA_FILE = path.join(LOCAL_DATA_DIR, 'mibi-dev.json');

const buildDefaultLocalUsers = (): LocalUsuarioRecord[] => {
  const now = Date.now();
  const adminPass = process.env.ADMIN_INITIAL_PASSWORD || 'mibi-admin-dev';
  const ventasPass = process.env.VENTAS_INITIAL_PASSWORD || 'mibi-ventas-dev';
  const adminHash = hashPassword(adminPass);
  const ventasHash = hashPassword(ventasPass);

  return [
    {
      id: '00000000-0000-4000-8000-000000000001',
      username: 'admin',
      nombre: 'Administración',
      rol: 'ADMIN',
      passwordHash: adminHash,
      activo: true,
      tokenVersion: 1,
      createdAt: now,
      updatedAt: now
    },
    {
      id: '00000000-0000-4000-8000-000000000002',
      username: 'juancho',
      nombre: 'Juancho',
      rol: 'VENTAS_DEPOSITO',
      passwordHash: ventasHash,
      activo: true,
      tokenVersion: 1,
      createdAt: now,
      updatedAt: now
    },
    {
      id: '00000000-0000-4000-8000-000000000003',
      username: 'pablo',
      nombre: 'Pablo Palacios',
      rol: 'VENTAS_DEPOSITO',
      passwordHash: ventasHash,
      activo: true,
      tokenVersion: 1,
      createdAt: now,
      updatedAt: now
    },
    {
      id: '00000000-0000-4000-8000-000000000004',
      username: 'agustin',
      nombre: 'Agustín',
      rol: 'VENTAS_DEPOSITO',
      passwordHash: ventasHash,
      activo: true,
      tokenVersion: 1,
      createdAt: now,
      updatedAt: now
    },
    {
      id: '00000000-0000-4000-8000-000000000005',
      username: 'rocio',
      nombre: 'Rocío',
      rol: 'VENTAS_DEPOSITO',
      passwordHash: ventasHash,
      activo: true,
      tokenVersion: 1,
      createdAt: now,
      updatedAt: now
    },
    {
      id: '00000000-0000-4000-8000-000000000006',
      username: 'deposito',
      nombre: 'Depósito',
      rol: 'DEPOSITO',
      passwordHash: ventasHash,
      activo: true,
      tokenVersion: 1,
      createdAt: now,
      updatedAt: now
    }
  ];
};

export const readLocalStore = (): LocalDevStore => {
  if (!fs.existsSync(LOCAL_DATA_DIR)) {
    fs.mkdirSync(LOCAL_DATA_DIR, { recursive: true });
  }
  if (!fs.existsSync(LOCAL_DATA_FILE)) {
    const initial: LocalDevStore = {
      nextNumeroSeq: 0,
      nextCustomCodeSeq: 0,
      usuarios: buildDefaultLocalUsers(),
      pedidos: [],
      productosCustom: [],
      pedidoAsignaciones: []
    };
    fs.writeFileSync(LOCAL_DATA_FILE, JSON.stringify(initial, null, 2), 'utf-8');
    return initial;
  }

  try {
    const raw = fs.readFileSync(LOCAL_DATA_FILE, 'utf-8');
    const parsed = JSON.parse(raw) as Partial<LocalDevStore>;
    const usuarios = Array.isArray(parsed.usuarios) && parsed.usuarios.length > 0
      ? parsed.usuarios
      : buildDefaultLocalUsers();
    const pedidos = Array.isArray(parsed.pedidos) ? parsed.pedidos : [];
    const productosCustom = Array.isArray(parsed.productosCustom) ? parsed.productosCustom : [];
    const pedidoAsignaciones = Array.isArray(parsed.pedidoAsignaciones) ? parsed.pedidoAsignaciones : [];
    const nextNumeroSeq = typeof parsed.nextNumeroSeq === 'number' ? parsed.nextNumeroSeq : 0;
    const nextCustomCodeSeq = typeof parsed.nextCustomCodeSeq === 'number' ? parsed.nextCustomCodeSeq : 0;
    return {
      nextNumeroSeq,
      nextCustomCodeSeq,
      usuarios,
      pedidos,
      productosCustom,
      pedidoAsignaciones
    };
  } catch {
    const fallback: LocalDevStore = {
      nextNumeroSeq: 0,
      nextCustomCodeSeq: 0,
      usuarios: buildDefaultLocalUsers(),
      pedidos: [],
      productosCustom: [],
      pedidoAsignaciones: []
    };
    fs.writeFileSync(LOCAL_DATA_FILE, JSON.stringify(fallback, null, 2), 'utf-8');
    return fallback;
  }
};

export const writeLocalStore = (store: LocalDevStore): void => {
  if (!fs.existsSync(LOCAL_DATA_DIR)) {
    fs.mkdirSync(LOCAL_DATA_DIR, { recursive: true });
  }
  const tmpFile = `${LOCAL_DATA_FILE}.tmp`;
  fs.writeFileSync(tmpFile, JSON.stringify(store, null, 2), 'utf-8');
  fs.renameSync(tmpFile, LOCAL_DATA_FILE);
};
