'use client';

import React, { useEffect, useState } from 'react';
import { CheckCircle2, KeyRound, ShieldCheck, UserPlus, X } from 'lucide-react';
import {
  cambiarRolUsuarioAction,
  crearUsuarioAction,
  listarUsuariosAction,
  rotarPasswordAction,
  toggleActivoUsuarioAction
} from '@/app/actions/usuarios-actions';
import { ejecutarSuiteQA } from '@/lib/qa-suite';
import { rolUsuarioToLabel, type ResultadoQA, type RolUsuario, type SessionUser, type ToneStatus, type UsuarioEquipo } from '@/lib/types';
import { StatusPill } from './status-pill';

interface EquipoModalProps {
  session: SessionUser;
  onClose: () => void;
  onNotify: (texto: string, tone?: ToneStatus) => void;
  onEquipoActualizado?: () => Promise<void>;
}

export function EquipoModal({
  session,
  onClose,
  onNotify,
  onEquipoActualizado
}: EquipoModalProps): React.JSX.Element {
  const esAdmin = session.rol === 'ADMIN';
  const [usuarios, setUsuarios] = useState<UsuarioEquipo[]>([]);
  const [cargando, setCargando] = useState<boolean>(esAdmin);
  const [nuevoUsername, setNuevoUsername] = useState<string>('');
  const [nuevoNombre, setNuevoNombre] = useState<string>('');
  const [nuevoRol, setNuevoRol] = useState<RolUsuario>('VENTAS');
  const [nuevaPassword, setNuevaPassword] = useState<string>('');
  const [guardando, setGuardando] = useState<boolean>(false);
  const [rotandoId, setRotandoId] = useState<string | null>(null);
  const [passwordRotada, setPasswordRotada] = useState<string>('');
  const [resultadosVerificacion, setResultadosVerificacion] = useState<ResultadoQA[] | null>(null);

  const cargarUsuarios = async (): Promise<void> => {
    if (!esAdmin) return;
    setCargando(true);
    const res = await listarUsuariosAction();
    if (res.ok && res.data) {
      setUsuarios(res.data);
    }
    setCargando(false);
  };

  useEffect(() => {
    void cargarUsuarios();
  }, [esAdmin]);

  const handleCrearUsuario = async (e: React.FormEvent<HTMLFormElement>): Promise<void> => {
    e.preventDefault();
    setGuardando(true);
    const res = await crearUsuarioAction({
      username: nuevoUsername.trim(),
      nombre: nuevoNombre.trim(),
      rol: nuevoRol,
      password: nuevaPassword
    });
    setGuardando(false);
    if (!res.ok) {
      onNotify(res.error ?? 'No se pudo crear el usuario.', 'red');
      return;
    }
    setNuevoUsername('');
    setNuevoNombre('');
    setNuevaPassword('');
    onNotify(`Usuario ${res.data?.nombre} creado.`, 'green');
    await cargarUsuarios();
    if (onEquipoActualizado) await onEquipoActualizado();
  };

  const handleCambiarRol = async (u: UsuarioEquipo, rol: RolUsuario): Promise<void> => {
    const res = await cambiarRolUsuarioAction({ userId: u.id, rol });
    if (!res.ok) {
      onNotify(res.error ?? 'No se pudo cambiar el rol.', 'red');
      return;
    }
    onNotify(`Rol de ${u.nombre} actualizado a ${rolUsuarioToLabel(rol)}.`, 'green');
    await cargarUsuarios();
    if (onEquipoActualizado) await onEquipoActualizado();
  };

  const handleRotarPassword = async (userId: string): Promise<void> => {
    if (passwordRotada.trim().length < 4) {
      onNotify('La nueva contraseña debe tener mínimo 4 caracteres.', 'amber');
      return;
    }
    setGuardando(true);
    const res = await rotarPasswordAction({ userId, nuevaPassword: passwordRotada });
    setGuardando(false);
    if (!res.ok) {
      onNotify(res.error ?? 'No se pudo cambiar la contraseña.', 'red');
      return;
    }
    setRotandoId(null);
    setPasswordRotada('');
    onNotify('Contraseña actualizada. Las sesiones anteriores de ese usuario fueron cerradas.', 'green');
    await cargarUsuarios();
  };

  const handleToggleActivo = async (u: UsuarioEquipo): Promise<void> => {
    const res = await toggleActivoUsuarioAction({ userId: u.id, activo: !u.activo });
    if (!res.ok) {
      onNotify(res.error ?? 'No se pudo actualizar el estado.', 'red');
      return;
    }
    onNotify(u.activo ? `${u.nombre} deshabilitado.` : `${u.nombre} habilitado.`, 'neutral');
    await cargarUsuarios();
    if (onEquipoActualizado) await onEquipoActualizado();
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/30 px-4"
      onClick={onClose}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        className="w-full max-w-xl bg-white rounded-2xl border border-[#EAEAEA] shadow-2xl p-6 space-y-5 max-h-[88vh] overflow-y-auto"
      >
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-4 h-4 text-zinc-700" strokeWidth={1.5} />
            <h2 className="text-sm font-semibold text-[#111111]">Equipo y verificación del sistema</h2>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-900"
          >
            <X className="w-4 h-4" strokeWidth={1.5} />
          </button>
        </div>

        <div className="bg-[#FAFAFA] rounded-xl border border-zinc-200/80 px-4 py-3 text-xs flex items-center justify-between">
          <div>
            <div className="font-semibold text-[#111111]">
              {session.nombre} <span className="text-zinc-400 font-normal">(@{session.username})</span>
            </div>
            <div className="text-zinc-500">Perfil activo: {rolUsuarioToLabel(session.rol)}</div>
          </div>
          <StatusPill tone={esAdmin ? 'green' : 'neutral'}>{rolUsuarioToLabel(session.rol)}</StatusPill>
        </div>

        {esAdmin && (
          <div className="space-y-4 pt-2 border-t border-zinc-200/70">
            <div className="text-xs font-semibold text-zinc-600">Usuarios internos</div>

            {cargando ? (
              <div className="text-xs text-zinc-400 py-3">Cargando equipo...</div>
            ) : (
              <div className="divide-y divide-zinc-200/60 border border-zinc-200/80 rounded-xl">
                {usuarios.map((u) => (
                  <div key={u.id} className="px-3.5 py-2.5 text-xs space-y-2">
                    <div className="flex items-center justify-between gap-2">
                      <div>
                        <span className="font-semibold text-[#111111]">{u.nombre}</span>{' '}
                        <span className="text-zinc-400">@{u.username}</span>
                        <div className="text-[11px] text-zinc-500">{rolUsuarioToLabel(u.rol)}</div>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <StatusPill tone={u.activo ? 'green' : 'red'}>
                          {u.activo ? 'Activo' : 'Inactivo'}
                        </StatusPill>
                        {u.id !== session.userId && (
                          <select
                            value={u.rol}
                            onChange={(e) => void handleCambiarRol(u, e.target.value as RolUsuario)}
                            className="px-2 py-1 rounded-full border border-zinc-200 bg-white text-[11px] text-zinc-700"
                            aria-label={`Rol de ${u.nombre}`}
                          >
                            <option value="VENTAS">Ventas</option>
                            <option value="DEPOSITO">Depósito</option>
                            <option value="ADMIN">Administración</option>
                            {u.rol === 'VENTAS_DEPOSITO' && (
                              <option value="VENTAS_DEPOSITO">Ventas / Depósito (legado)</option>
                            )}
                          </select>
                        )}
                        <button
                          type="button"
                          onClick={() => {
                            setRotandoId(rotandoId === u.id ? null : u.id);
                            setPasswordRotada('');
                          }}
                          className="px-2.5 py-1 rounded-full border border-zinc-200 text-[11px] text-zinc-700 hover:bg-zinc-50 inline-flex items-center gap-1"
                        >
                          <KeyRound className="w-3 h-3" strokeWidth={1.5} />
                          Clave
                        </button>
                        {u.id !== session.userId && (
                          <button
                            type="button"
                            onClick={() => void handleToggleActivo(u)}
                            className="px-2.5 py-1 rounded-full border border-zinc-200 text-[11px] text-zinc-700 hover:bg-zinc-50"
                          >
                            {u.activo ? 'Desactivar' : 'Activar'}
                          </button>
                        )}
                      </div>
                    </div>

                    {rotandoId === u.id && (
                      <div className="flex items-center gap-2 pt-1">
                        <input
                          type="password"
                          value={passwordRotada}
                          onChange={(e) => setPasswordRotada(e.target.value)}
                          placeholder="Nueva contraseña (mín. 4 caracteres)"
                          className="flex-1 px-3 py-1.5 text-xs bg-white border border-zinc-200 rounded-lg"
                        />
                        <button
                          type="button"
                          disabled={guardando}
                          onClick={() => void handleRotarPassword(u.id)}
                          className="px-3 py-1.5 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800"
                        >
                          Guardar
                        </button>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}

            <form onSubmit={(e) => void handleCrearUsuario(e)} className="space-y-3 pt-2 border-t border-zinc-200/60">
              <div className="text-xs font-semibold text-zinc-600 flex items-center gap-1.5">
                <UserPlus className="w-3.5 h-3.5" strokeWidth={1.5} />
                Alta de usuario interno
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 text-xs">
                <input
                  type="text"
                  value={nuevoUsername}
                  onChange={(e) => setNuevoUsername(e.target.value)}
                  placeholder="Usuario (ej. lucas)"
                  className="px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                />
                <input
                  type="text"
                  value={nuevoNombre}
                  onChange={(e) => setNuevoNombre(e.target.value)}
                  placeholder="Nombre visible (ej. Lucas)"
                  className="px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                />
                <select
                  value={nuevoRol}
                  onChange={(e) => setNuevoRol(e.target.value as RolUsuario)}
                  className="px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                >
                  <option value="VENTAS">Ventas</option>
                  <option value="DEPOSITO">Depósito</option>
                  <option value="ADMIN">Administración</option>
                </select>
                <input
                  type="password"
                  value={nuevaPassword}
                  onChange={(e) => setNuevaPassword(e.target.value)}
                  placeholder="Contraseña inicial"
                  className="px-3 py-2 bg-white border border-zinc-200 rounded-xl"
                />
              </div>
              <button
                type="submit"
                disabled={guardando}
                className="px-4 py-2 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800 disabled:opacity-50"
              >
                Crear usuario
              </button>
            </form>
          </div>
        )}

        <div className="pt-4 border-t border-zinc-200/70 space-y-2.5">
          <button
            type="button"
            onClick={() => setResultadosVerificacion(ejecutarSuiteQA())}
            className="w-full inline-flex items-center justify-center gap-1.5 py-2 rounded-full text-xs font-medium bg-[#FAFAFA] border border-zinc-200 text-zinc-700 hover:bg-zinc-100 transition-colors"
          >
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" strokeWidth={1.5} />
            Ejecutar verificación técnica del motor (30 pruebas)
          </button>

          {resultadosVerificacion && (
            <div className="space-y-1.5 max-h-60 overflow-y-auto border border-zinc-200/80 rounded-xl p-3 bg-[#FAFAFA]">
              <div className="text-xs font-semibold text-[#111111] pb-1 border-b border-zinc-200/70 tabular-nums">
                Resultado: {resultadosVerificacion.filter((r) => r.ok).length} / {resultadosVerificacion.length}{' '}
                aprobadas
              </div>
              {resultadosVerificacion.map((r, idx) => (
                <div key={idx} className="flex items-start justify-between gap-2 text-[11px]">
                  <span className={r.ok ? 'text-zinc-700' : 'text-red-700 font-semibold'}>
                    {r.ok ? '✓' : '✗'} {r.nombre}
                  </span>
                  <span className="text-zinc-400 shrink-0 tabular-nums">{r.detalle}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
