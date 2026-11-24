import { randomBytes, scryptSync } from 'node:crypto';
import postgres from 'postgres';
import { drizzle } from 'drizzle-orm/postgres-js';
import { eq } from 'drizzle-orm';
import * as schema from '../src/lib/db/schema';

const hashPasswordLocal = (plainPassword: string): string => {
  const salt = randomBytes(16).toString('hex');
  const derived = scryptSync(plainPassword.trim(), salt, 64).toString('hex');
  return `scrypt$v1$${salt}$${derived}`;
};

async function main(): Promise<void> {
  const dbUrl = process.env.DIRECT_URL || process.env.DATABASE_URL;
  const adminPassword = process.env.ADMIN_INITIAL_PASSWORD?.trim();
  const ventasPassword = process.env.VENTAS_INITIAL_PASSWORD?.trim();

  if (!dbUrl) {
    throw new Error('Falta DIRECT_URL o DATABASE_URL para ejecutar el seed.');
  }
  if (!adminPassword || adminPassword.length < 4) {
    throw new Error('Falta ADMIN_INITIAL_PASSWORD (mínimo 4 caracteres).');
  }
  if (!ventasPassword || ventasPassword.length < 4) {
    throw new Error('Falta VENTAS_INITIAL_PASSWORD (mínimo 4 caracteres).');
  }

  const client = postgres(dbUrl, { prepare: false, ssl: 'require', max: 1 });
  const db = drizzle(client, { schema });

  const seedUsers: Array<{ username: string; nombre: string; rol: 'ADMIN' | 'VENTAS_DEPOSITO'; password: string }> = [
    { username: 'admin', nombre: 'Administración', rol: 'ADMIN', password: adminPassword },
    { username: 'juancho', nombre: 'Juancho', rol: 'VENTAS_DEPOSITO', password: ventasPassword },
    { username: 'pablo', nombre: 'Pablo Palacios', rol: 'VENTAS_DEPOSITO', password: ventasPassword },
    { username: 'agustin', nombre: 'Agustín', rol: 'VENTAS_DEPOSITO', password: ventasPassword },
    { username: 'rocio', nombre: 'Rocío', rol: 'VENTAS_DEPOSITO', password: ventasPassword },
    { username: 'deposito', nombre: 'Depósito', rol: 'VENTAS_DEPOSITO', password: ventasPassword }
  ];

  for (const item of seedUsers) {
    const existing = await db
      .select({ id: schema.usuarios.id, tokenVersion: schema.usuarios.tokenVersion })
      .from(schema.usuarios)
      .where(eq(schema.usuarios.username, item.username))
      .limit(1);

    const passwordHash = hashPasswordLocal(item.password);
    if (existing[0]) {
      await db
        .update(schema.usuarios)
        .set({
          nombre: item.nombre,
          rol: item.rol,
          passwordHash,
          activo: true,
          tokenVersion: existing[0].tokenVersion + 1,
          updatedAt: new Date()
        })
        .where(eq(schema.usuarios.id, existing[0].id));
      console.log(`Actualizado: ${item.username} (${item.rol})`);
    } else {
      await db.insert(schema.usuarios).values({
        username: item.username,
        nombre: item.nombre,
        rol: item.rol,
        passwordHash,
        activo: true,
        tokenVersion: 1
      });
      console.log(`Creado: ${item.username} (${item.rol})`);
    }
  }

  await client.end();
}

main().catch((err: unknown) => {
  console.error(err instanceof Error ? err.message : err);
  process.exit(1);
});
