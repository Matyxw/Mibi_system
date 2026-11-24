'use client';

import React, { useState } from 'react';
import { useRouter } from 'next/navigation';
import { MIBI_BRAND } from '@/lib/catalogo-mibi';
import { loginAction } from '@/app/actions/auth-actions';
import { MibiLogo } from './mibi-logo';

export function LoginForm({ isLocalDev }: { isLocalDev: boolean }): React.JSX.Element {
  const router = useRouter();
  const [username, setUsername] = useState<string>('');
  const [password, setPassword] = useState<string>('');
  const [error, setError] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>): Promise<void> => {
    e.preventDefault();
    setError('');
    if (!username.trim() || !password) {
      setError('Completá tu usuario y contraseña.');
      return;
    }

    setLoading(true);
    try {
      const res = await loginAction({ username: username.trim(), password });
      if (!res.ok) {
        setError(res.error ?? 'Usuario o contraseña incorrectos');
        return;
      }
      router.push('/');
      router.refresh();
    } catch {
      setError('No se pudo iniciar sesión. Intentá nuevamente.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#FAFAFA] text-[#111111] flex flex-col items-center justify-center px-4 py-12">
      <div className="h-[2px] w-full bg-[#F4B6CE] fixed top-0 left-0 z-10" />

      <div className="w-full max-w-sm bg-white rounded-2xl border border-[#EAEAEA] shadow-[0_1px_2px_rgba(0,0,0,0.03)] p-7 space-y-6">
        <div className="flex flex-col items-center text-center">
          <div className="pt-2 pb-1 flex flex-col items-center justify-center">
            <MibiLogo width={108} variant="pink" />
            <span className="mt-2.5 inline-flex items-center px-3 py-0.5 rounded-full text-[11px] font-medium bg-[#FCE7F3] text-[#111111] border border-[#F4B8D1] tracking-tight">
              Gestión Mayorista
            </span>
          </div>

          <h1 className="mt-5 text-base font-semibold tracking-tight text-[#111111]">
            Ingreso al equipo
          </h1>
          <p className="mt-1 text-xs text-zinc-500">
            Ingresá con tu usuario interno y contraseña asignada.
          </p>
          {isLocalDev && (
            <span className="mt-2 inline-flex items-center px-2.5 py-0.5 rounded-full text-[10px] font-medium bg-zinc-100 text-zinc-600 border border-zinc-200">
              Entorno Local
            </span>
          )}
        </div>

        <form onSubmit={(e) => void handleSubmit(e)} className="space-y-4">
          <div>
            <label className="block text-xs font-medium text-zinc-600 mb-1.5">Usuario</label>
            <input
              type="text"
              autoComplete="username"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="Ej. admin, juancho, rocio..."
              autoFocus
              className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 focus:ring-2 focus:ring-pink-200/60 transition-all"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-zinc-600 mb-1.5">Contraseña</label>
            <input
              type="password"
              autoComplete="current-password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Ingresá tu contraseña"
              className="w-full px-3.5 py-2.5 text-sm bg-white border border-zinc-200 rounded-xl focus:outline-none focus:border-zinc-900 focus:ring-2 focus:ring-pink-200/60 transition-all"
            />
          </div>

          {error && <p className="text-xs text-red-600">{error}</p>}

          <button
            type="submit"
            disabled={loading}
            className="w-full py-2.5 rounded-full text-sm font-medium bg-zinc-900 hover:bg-zinc-800 text-white transition-colors disabled:opacity-50"
          >
            {loading ? 'Verificando...' : 'Ingresar'}
          </button>
        </form>
      </div>

      <p className="mt-6 text-[11px] text-zinc-400 text-center">{MIBI_BRAND.pieOficial}</p>
    </div>
  );
}
