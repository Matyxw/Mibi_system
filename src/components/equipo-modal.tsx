'use client';

import React, { useEffect, useState } from 'react';
import { AlertTriangle, CheckCircle2, ChevronDown, KeyRound, RotateCcw, ShieldCheck, UserPlus, X } from 'lucide-react';
import {
  cambiarRolUsuarioAction,
  crearUsuarioAction,
  listarUsuariosAction,
  rotarPasswordAction,
  toggleActivoUsuarioAction
} from '@/app/actions/usuarios-actions';
import { ejecutarSuiteQA } from '@/lib/qa-suite';
import {
  rolUsuarioToLabel,
  type ResultadoQA,
  type RolUsuario,
  type SessionUser,
  type ToneStatus,
  type UsuarioEquipo
} from '@/lib/types';
import { StatusPill } from './status-pill';

interface EquipoModalProps {
  session: SessionUser;
  onClose: () => void;
  onNotify: (texto: string, tone?: ToneStatus) => void;
  onEquipoActualizado?: () => Promise<void>;
  onReiniciarSistema?: () => Promise<boolean>;
}

export function EquipoModal({
  session,
  onClose,
  onNotify,
  onEquipoActualizado,
  onReiniciarSistema
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

  const [mostrarQA, setMostrarQA] = useState<boolean>(false);
  const [resultadosVerificacion, setResultadosVerificacion] = useState<ResultadoQA[] | null>(null);

  const [mostrarPeligro, setMostrarPeligro] = useState<boolean>(false);
  const [confirmTextoReinicio, setConfirmTextoReinicio] = useState<string>('');
  const [reiniciando, setReiniciando] = useState<boolean>(false);

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
    if (guardando) return;
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
    if (guardando) return;
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

  const handleConfirmarReinicio = async (): Promise<void> => {
    if (!onReiniciarSistema || confirmTextoReinicio.trim().toUpperCase() !== 'REINICIAR' || reiniciando) {
      return;
    }
    setReiniciando(true);
    const ok = await onReiniciarSistema();
    setReiniciando(false);
    if (ok) {
      setConfirmTextoReinicio('');
      setMostrarPeligro(false);
      onClose();
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/30 dark:bg-black/65 px-4"
      onClick={onClose}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        className="w-full max-w-xl bg-white dark:bg-[#141417] rounded-2xl border border-[#EAEAEA] dark:border-[#26262B] shadow-2xl p-6 space-y-5 max-h-[88vh] overflow-y-auto"
      >
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-4 h-4 text-zinc-700 dark:text-zinc-300" strokeWidth={1.5} />
            <h2 className="text-sm font-semibold text-[#111111] dark:text-[#F4F4F5]">
              {esAdmin ? 'Equipo y configuración' : 'Mi cuenta'}
            </h2>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-full text-zinc-400 hover:text-zinc-900 dark:hover:text-[#F4F4F5]"
          >
            <X className="w-4 h-4" strokeWidth={1.5} />
          </button>
        </div>

        <div className="bg-[#FAFAFA] dark:bg-[#1B1B1F] rounded-xl border border-zinc-200/80 dark:border-[#26262B] px-4 py-3 text-xs flex items-center justify-between">
          <div>
            <div className="font-semibold text-[#111111] dark:text-[#F4F4F5]">
              {session.nombre}{' '}
              <span className="text-zinc-400 dark:text-zinc-500 font-normal">(@{session.username})</span>
            </div>
            <div className="text-zinc-500 dark:text-zinc-400">Perfil activo: {rolUsuarioToLabel(session.rol)}</div>
          </div>
          <StatusPill tone={esAdmin ? 'green' : 'neutral'}>{rolUsuarioToLabel(session.rol)}</StatusPill>
        </div>

        {esAdmin && (
          <div className="space-y-4 pt-2 border-t border-zinc-200/70 dark:border-[#26262B]">
            <div className="text-xs font-semibold text-zinc-600 dark:text-zinc-300">Usuarios internos</div>

            {cargando ? (
              <div className="text-xs text-zinc-400 dark:text-zinc-500 py-3">Cargando equipo...</div>
            ) : (
              <div className="divide-y divide-zinc-200/60 dark:divide-[#26262B]/80 border border-zinc-200/80 dark:border-[#26262B] rounded-xl">
                {usuarios.map((u) => {
                  const esYo = u.id === session.userId;
                  return (
                    <div key={u.id} className="px-3.5 py-2.5 text-xs space-y-2">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div>
                          <span className="font-semibold text-[#111111] dark:text-[#F4F4F5]">{u.nombre}</span>{' '}
                          <span className="text-zinc-400 dark:text-zinc-500">@{u.username}</span>
                          <div className="text-[11px] text-zinc-500 dark:text-zinc-400">
                            {rolUsuarioToLabel(u.rol)}
                            {esYo ? ' · Tu sesión' : ''}
                          </div>
                        </div>
                        <div className="flex items-center gap-1.5 flex-wrap">
                          <StatusPill tone={u.activo ? 'green' : 'red'}>
                            {u.activo ? 'Activo' : 'Inactivo'}
                          </StatusPill>
                          {!esYo && (
                            <select
                              value={u.rol}
                              onChange={(e) => void handleCambiarRol(u, e.target.value as RolUsuario)}
                              className="px-2 py-1 rounded-full border border-zinc-200 dark:border-[#2E2E35] bg-white dark:bg-[#1B1B1F] text-[11px] text-zinc-700 dark:text-zinc-200"
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
                            className="px-2.5 py-1 rounded-full border border-zinc-200 dark:border-[#2E2E35] dark:bg-[#1B1B1F] text-[11px] text-zinc-700 dark:text-zinc-200 hover:bg-zinc-50 dark:hover:bg-[#24242B] inline-flex items-center gap-1"
                          >
                            <KeyRound className="w-3 h-3" strokeWidth={1.5} />
                            Clave
                          </button>
                          {!esYo && (
                            <button
                              type="button"
                              onClick={() => void handleToggleActivo(u)}
                              className="px-2.5 py-1 rounded-full border border-zinc-200 dark:border-[#2E2E35] dark:bg-[#1B1B1F] text-[11px] text-zinc-700 dark:text-zinc-200 hover:bg-zinc-50 dark:hover:bg-[#24242B]"
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
                            className="flex-1 px-3 py-1.5 text-xs bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-lg placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                          />
                          <button
                            type="button"
                            disabled={guardando}
                            onClick={() => void handleRotarPassword(u.id)}
                            className="px-3 py-1.5 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800 dark:bg-zinc-100 dark:text-zinc-950 dark:hover:bg-white"
                          >
                            Guardar
                          </button>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}

            <form
              onSubmit={(e) => void handleCrearUsuario(e)}
              className="space-y-3 pt-2 border-t border-zinc-200/60 dark:border-[#26262B]"
            >
              <div className="text-xs font-semibold text-zinc-600 dark:text-zinc-300 flex items-center gap-1.5">
                <UserPlus className="w-3.5 h-3.5" strokeWidth={1.5} />
                Alta de usuario interno
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 text-xs">
                <input
                  type="text"
                  value={nuevoUsername}
                  onChange={(e) => setNuevoUsername(e.target.value)}
                  placeholder="Usuario (ej. lucas)"
                  className="px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                />
                <input
                  type="text"
                  value={nuevoNombre}
                  onChange={(e) => setNuevoNombre(e.target.value)}
                  placeholder="Nombre visible (ej. Lucas)"
                  className="px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                />
                <select
                  value={nuevoRol}
                  onChange={(e) => setNuevoRol(e.target.value as RolUsuario)}
                  className="px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl"
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
                  className="px-3 py-2 bg-white dark:bg-[#1B1B1F] dark:text-[#F4F4F5] border border-zinc-200 dark:border-[#2E2E35] rounded-xl placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                />
              </div>
              <button
                type="submit"
                disabled={guardando}
                className="px-4 py-2 rounded-full text-xs font-medium bg-zinc-900 text-white hover:bg-zinc-800 dark:bg-zinc-100 dark:text-zinc-950 dark:hover:bg-white disabled:opacity-50"
              >
                Crear usuario
              </button>
            </form>
          </div>
        )}

        <div className="pt-4 border-t border-zinc-200/70 dark:border-[#26262B] space-y-2.5">
          <button
            type="button"
            onClick={() => {
              const next = !mostrarQA;
              setMostrarQA(next);
              if (next && !resultadosVerificacion) {
                setResultadosVerificacion(ejecutarSuiteQA());
              }
            }}
            className="w-full flex items-center justify-between px-3.5 py-2 rounded-xl text-xs font-medium bg-[#FAFAFA] dark:bg-[#1B1B1F] border border-zinc-200/80 dark:border-[#26262B] text-zinc-700 dark:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-[#24242B] transition-colors"
          >
            <span className="inline-flex items-center gap-1.5">
              <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400" strokeWidth={1.5} />
              Verificación técnica del sistema
            </span>
            <ChevronDown
              className={`w-3.5 h-3.5 text-zinc-400 transition-transform ${mostrarQA ? 'rotate-180' : ''}`}
              strokeWidth={1.5}
            />
          </button>

          {mostrarQA && resultadosVerificacion && (
            <div className="space-y-2 border border-zinc-200/80 dark:border-[#26262B] rounded-xl p-3 bg-[#FAFAFA] dark:bg-[#1B1B1F]">
              <div className="flex items-center justify-between pb-1 border-b border-zinc-200/70 dark:border-[#26262B]">
                <span className="text-xs font-semibold text-[#111111] dark:text-[#F4F4F5] tabular-nums">
                  Resultado: {resultadosVerificacion.filter((r) => r.ok).length} / {resultadosVerificacion.length}{' '}
                  aprobadas
                </span>
                <button
                  type="button"
                  onClick={() => setResultadosVerificacion(ejecutarSuiteQA())}
                  className="text-[11px] font-medium text-zinc-600 dark:text-zinc-300 hover:underline"
                >
                  Re-ejecutar
                </button>
              </div>
              <div className="space-y-1.5 max-h-52 overflow-y-auto">
                {resultadosVerificacion.map((r, idx) => (
                  <div key={idx} className="flex items-start justify-between gap-2 text-[11px]">
                    <span
                      className={
                        r.ok ? 'text-zinc-700 dark:text-zinc-300' : 'text-red-700 dark:text-red-400 font-semibold'
                      }
                    >
                      {r.ok ? '✓' : '✗'} {r.nombre}
                    </span>
                    <span className="text-zinc-400 dark:text-zinc-500 shrink-0 tabular-nums">{r.detalle}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {esAdmin && onReiniciarSistema && (
            <div className="pt-1">
              <button
                type="button"
                onClick={() => {
                  setMostrarPeligro((prev) => !prev);
                  setConfirmTextoReinicio('');
                }}
                className="w-full flex items-center justify-between px-3.5 py-2 rounded-xl text-xs font-medium border border-red-200/70 dark:border-red-500/25 text-red-700 dark:text-red-400 hover:bg-red-50/70 dark:hover:bg-red-500/10 transition-colors"
              >
                <span className="inline-flex items-center gap-1.5">
                  <AlertTriangle className="w-3.5 h-3.5" strokeWidth={1.5} />
                  Zona de peligro (reiniciar pedidos)
                </span>
                <ChevronDown
                  className={`w-3.5 h-3.5 transition-transform ${mostrarPeligro ? 'rotate-180' : ''}`}
                  strokeWidth={1.5}
                />
              </button>

              {mostrarPeligro && (
                <div className="mt-2 bg-red-50/70 dark:bg-red-500/10 border border-red-200 dark:border-red-500/25 rounded-xl p-3.5 space-y-2.5 text-xs">
                  <p className="text-red-800 dark:text-red-200">
                    Elimina todos los pedidos y reinicia el numerador en #0001. Para confirmar, escribí{' '}
                    <strong>REINICIAR</strong> abajo:
                  </p>
                  <div className="flex items-center gap-2">
                    <input
                      type="text"
                      value={confirmTextoReinicio}
                      onChange={(e) => setConfirmTextoReinicio(e.target.value)}
                      placeholder="Escribí REINICIAR"
                      className="flex-1 px-3 py-1.5 bg-white dark:bg-[#141417] dark:text-[#F4F4F5] border border-red-200 dark:border-red-500/30 rounded-lg uppercase"
                    />
                    <button
                      type="button"
                      disabled={
                        reiniciando || confirmTextoReinicio.trim().toUpperCase() !== 'REINICIAR'
                      }
                      onClick={() => void handleConfirmarReinicio()}
                      className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-full bg-red-600 hover:bg-red-700 text-white font-medium disabled:opacity-40"
                    >
                      <RotateCcw className="w-3.5 h-3.5" strokeWidth={1.5} />
                      {reiniciando ? 'Reiniciando...' : 'Confirmar reinicio'}
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
