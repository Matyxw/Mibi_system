-- ============================================================================
-- MIBI® ERP MAYORISTA — ESQUEMA SUPABASE POSTGRESQL (MODO BÚNKER)
-- 100% Idempotente: ejecutar completo en Supabase SQL Editor
-- ============================================================================

CREATE EXTENSION IF NOT EXISTS pgcrypto;

-- 1. USUARIOS INTERNOS (SIN REGISTRO PÚBLICO)
CREATE TABLE IF NOT EXISTS public.usuarios (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  username TEXT UNIQUE NOT NULL,
  nombre TEXT NOT NULL,
  rol TEXT NOT NULL CHECK (rol IN ('ADMIN', 'VENTAS', 'DEPOSITO', 'VENTAS_DEPOSITO')),
  password_hash TEXT NOT NULL,
  activo BOOLEAN NOT NULL DEFAULT true,
  token_version INT NOT NULL DEFAULT 1,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Actualización idempotente de restricción de roles en bases ya existentes
DO $$
DECLARE
  cname TEXT;
BEGIN
  FOR cname IN
    SELECT con.conname
    FROM pg_constraint con
    JOIN pg_class rel ON rel.oid = con.conrelid
    JOIN pg_namespace nsp ON nsp.oid = rel.relnamespace
    WHERE nsp.nspname = 'public'
      AND rel.relname = 'usuarios'
      AND con.contype = 'c'
      AND pg_get_constraintdef(con.oid) LIKE '%rol%'
  LOOP
    EXECUTE format('ALTER TABLE public.usuarios DROP CONSTRAINT IF EXISTS %I', cname);
  END LOOP;
  ALTER TABLE public.usuarios
    ADD CONSTRAINT usuarios_rol_check CHECK (rol IN ('ADMIN', 'VENTAS', 'DEPOSITO', 'VENTAS_DEPOSITO'));
END $$;

-- 2. SECUENCIA ATÓMICA DE NÚMERO DE PEDIDO
CREATE SEQUENCE IF NOT EXISTS public.pedidos_numero_seq START 1;

-- 3. TABLA PRINCIPAL DE PEDIDOS
CREATE TABLE IF NOT EXISTS public.pedidos (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  numero_pedido TEXT UNIQUE NOT NULL,
  cliente TEXT NOT NULL,
  vendedor TEXT NOT NULL DEFAULT 'Administración',
  whatsapp TEXT NOT NULL DEFAULT '',
  fecha_pedido DATE NOT NULL,
  despacho_previsto DATE,
  fecha_despacho DATE,
  fecha_entrega DATE,
  descuento_porc NUMERIC(5,2) NOT NULL DEFAULT 0 CHECK (descuento_porc >= 0 AND descuento_porc <= 100),
  subtotal_neto NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (subtotal_neto >= 0),
  monto_descuento NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (monto_descuento >= 0),
  base_imponible NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (base_imponible >= 0),
  iva_21 NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (iva_21 >= 0),
  total_facturado NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (total_facturado >= 0),
  total_facturado_override NUMERIC(14,2) NULL CHECK (total_facturado_override IS NULL OR total_facturado_override >= 0),
  medio_pago TEXT NOT NULL DEFAULT 'Sin pagar' CHECK (medio_pago IN ('Transferencia / Efectivo', 'Cheque', 'Sin pagar')),
  monto_abonado NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (monto_abonado >= 0 AND monto_abonado <= total_facturado),
  monto_cheque NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (monto_cheque >= 0 AND monto_abonado + monto_cheque <= total_facturado),
  cheque_acreditado BOOLEAN NOT NULL DEFAULT false,
  fecha_recepcion_cheque DATE,
  fecha_acreditacion_cheque DATE,
  detalle_cheque TEXT NOT NULL DEFAULT '',
  saldo_pendiente NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (saldo_pendiente >= 0),
  deuda_sin_respaldo NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (deuda_sin_respaldo >= 0),
  estado_pago TEXT NOT NULL DEFAULT 'Falta pago' CHECK (estado_pago IN ('Pagado', 'Pendiente', 'Falta pago')),
  estado_armado TEXT NOT NULL DEFAULT 'Por armar' CHECK (estado_armado IN ('Por armar', 'En armado', 'Listo', 'Falta talle o color')),
  estado_envio TEXT NOT NULL DEFAULT 'Pendiente' CHECK (estado_envio IN ('Pendiente', 'Despachado', 'Entregado')),
  transporte TEXT NOT NULL DEFAULT '',
  nro_seguimiento TEXT NOT NULL DEFAULT '',
  observaciones TEXT NOT NULL DEFAULT '',
  detalle_manual_importado TEXT,
  created_by UUID NULL REFERENCES public.usuarios(id) ON DELETE SET NULL,
  incidencia_deposito TEXT NOT NULL DEFAULT '',
  incidencia_at TIMESTAMPTZ NULL,
  incidencia_por TEXT NOT NULL DEFAULT '',
  version INT NOT NULL DEFAULT 1,
  deleted_at TIMESTAMPTZ NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE public.pedidos ADD COLUMN IF NOT EXISTS created_by UUID NULL REFERENCES public.usuarios(id) ON DELETE SET NULL;
ALTER TABLE public.pedidos ADD COLUMN IF NOT EXISTS incidencia_deposito TEXT NOT NULL DEFAULT '';
ALTER TABLE public.pedidos ADD COLUMN IF NOT EXISTS incidencia_at TIMESTAMPTZ NULL;
ALTER TABLE public.pedidos ADD COLUMN IF NOT EXISTS incidencia_por TEXT NOT NULL DEFAULT '';

-- 4. LÍNEAS DESGLOSADAS DE PEDIDO
CREATE TABLE IF NOT EXISTS public.pedido_lineas (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  pedido_id UUID NOT NULL REFERENCES public.pedidos(id) ON DELETE CASCADE,
  id_linea TEXT NOT NULL,
  id_grupo TEXT NOT NULL,
  nombre_producto TEXT NOT NULL,
  rubro TEXT NOT NULL,
  precio_unitario_neto NUMERIC(14,2) NOT NULL CHECK (precio_unitario_neto >= 0),
  cantidad_total INT NOT NULL CHECK (cantidad_total > 0),
  nota_cliente TEXT NOT NULL DEFAULT '',
  estado TEXT NOT NULL CHECK (estado IN ('OK', 'REQUIERE_REVISION')),
  requiere_revision BOOLEAN NOT NULL DEFAULT false,
  motivo_revision TEXT,
  desglose_json JSONB NOT NULL DEFAULT '[]'::jsonb
);

-- 5. HISTORIAL DE AUDITORÍA (APPEND-ONLY)
CREATE TABLE IF NOT EXISTS public.pedido_historial (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  pedido_id UUID NOT NULL REFERENCES public.pedidos(id) ON DELETE CASCADE,
  usuario_id UUID,
  usuario_nombre TEXT NOT NULL,
  rol TEXT NOT NULL,
  accion TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- 5B. CATÁLOGO EXTENSIBLE (PRODUCTOS FUERA DE PLANILLA)
CREATE TABLE IF NOT EXISTS public.productos_custom (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  id_grupo TEXT UNIQUE NOT NULL,
  nombre TEXT NOT NULL,
  rubro TEXT NOT NULL,
  precio_neto NUMERIC(14,2) NOT NULL CHECK (precio_neto >= 0),
  variantes_json JSONB NOT NULL DEFAULT '[]'::jsonb,
  activo BOOLEAN NOT NULL DEFAULT true,
  created_by UUID NULL REFERENCES public.usuarios(id) ON DELETE SET NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- 5C. DERIVACIÓN Y LIBERACIÓN DE PEDIDOS POR CUENTA (N:M)
CREATE TABLE IF NOT EXISTS public.pedido_asignaciones (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  pedido_id UUID NOT NULL REFERENCES public.pedidos(id) ON DELETE CASCADE,
  usuario_id UUID NOT NULL REFERENCES public.usuarios(id) ON DELETE CASCADE,
  enviado_por UUID NULL REFERENCES public.usuarios(id) ON DELETE SET NULL,
  enviado_por_nombre TEXT NOT NULL DEFAULT 'Administración',
  nota_envio TEXT NOT NULL DEFAULT '',
  prioridad TEXT NOT NULL DEFAULT 'NORMAL' CHECK (prioridad IN ('NORMAL', 'URGENTE')),
  visto_at TIMESTAMPTZ NULL,
  requiere_relectura BOOLEAN NOT NULL DEFAULT false,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (pedido_id, usuario_id)
);

-- 6. FUNCIÓN DE SINCRONIZACIÓN DE SECUENCIA TRAS IMPORTACIONES
CREATE OR REPLACE FUNCTION public.sync_pedidos_numero_seq()
RETURNS BIGINT
LANGUAGE plpgsql
AS $$
DECLARE
  max_num BIGINT;
BEGIN
  SELECT COALESCE(MAX(NULLIF(regexp_replace(numero_pedido, '\D', '', 'g'), '')::BIGINT), 0)
    INTO max_num
    FROM public.pedidos;
  IF max_num < 1 THEN
    PERFORM setval('public.pedidos_numero_seq', 1, false);
    RETURN 0;
  ELSE
    PERFORM setval('public.pedidos_numero_seq', max_num, true);
    RETURN max_num;
  END IF;
END;
$$;

-- 7. ÍNDICES DE RENDIMIENTO
CREATE INDEX IF NOT EXISTS idx_pedidos_deleted_created ON public.pedidos (deleted_at, created_at DESC);
CREATE INDEX IF NOT EXISTS idx_pedidos_estado_pago ON public.pedidos (estado_pago);
CREATE INDEX IF NOT EXISTS idx_pedidos_estado_armado ON public.pedidos (estado_armado);
CREATE INDEX IF NOT EXISTS idx_pedidos_fecha_cheque ON public.pedidos (fecha_acreditacion_cheque);
CREATE INDEX IF NOT EXISTS idx_pedidos_numero ON public.pedidos (numero_pedido);
CREATE INDEX IF NOT EXISTS idx_pedido_lineas_pedido_id ON public.pedido_lineas (pedido_id);
CREATE INDEX IF NOT EXISTS idx_pedido_historial_pedido_id ON public.pedido_historial (pedido_id, created_at DESC);
CREATE INDEX IF NOT EXISTS idx_pedidos_created_by ON public.pedidos (created_by);
CREATE INDEX IF NOT EXISTS idx_productos_custom_activo ON public.productos_custom (activo, created_at DESC);
CREATE INDEX IF NOT EXISTS idx_pedido_asignaciones_usuario_id ON public.pedido_asignaciones (usuario_id);
CREATE INDEX IF NOT EXISTS idx_pedido_asignaciones_pedido_id ON public.pedido_asignaciones (pedido_id);

-- 8. LOCKDOWN TOTAL DE LA API PÚBLICA DE SUPABASE (MODO BÚNKER)
ALTER TABLE public.usuarios ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.usuarios FORCE ROW LEVEL SECURITY;
ALTER TABLE public.pedidos ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pedidos FORCE ROW LEVEL SECURITY;
ALTER TABLE public.pedido_lineas ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pedido_lineas FORCE ROW LEVEL SECURITY;
ALTER TABLE public.pedido_historial ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pedido_historial FORCE ROW LEVEL SECURITY;
ALTER TABLE public.productos_custom ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.productos_custom FORCE ROW LEVEL SECURITY;
ALTER TABLE public.pedido_asignaciones ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pedido_asignaciones FORCE ROW LEVEL SECURITY;

REVOKE ALL ON ALL TABLES IN SCHEMA public FROM PUBLIC, anon, authenticated;
REVOKE ALL ON ALL SEQUENCES IN SCHEMA public FROM PUBLIC, anon, authenticated;
REVOKE ALL ON ALL FUNCTIONS IN SCHEMA public FROM PUBLIC, anon, authenticated;
