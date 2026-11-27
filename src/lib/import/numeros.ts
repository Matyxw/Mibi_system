import { round2 } from '@/lib/finanzas';

export const parsearNumeroDocumento = (raw: string | number | null | undefined): number => {
  if (raw === null || raw === undefined) return 0;
  if (typeof raw === 'number') return Number.isFinite(raw) ? round2(raw) : 0;

  const trimmed = raw
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, ' ')
    .replace(/\$/g, '')
    .trim();
  if (!trimmed) return 0;

  const token = trimmed.match(/^-?[\d.,]+/)?.[0] ?? trimmed.replace(/[^0-9.,-]/g, '');
  if (!token) return 0;

  const isNeg = token.startsWith('-');
  const abs = isNeg ? token.slice(1) : token;
  if (!/\d/.test(abs)) return 0;

  let val = 0;
  // Formato US explícito: miles con coma (23,800.00 / 830,408.04) o entero/decimal con punto y 2 decimales (901479.00)
  if (/^\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(abs) || /^\d+\.\d{1,2}$/.test(abs)) {
    val = parseFloat(abs.replace(/,/g, ''));
  } else if (/^\d+(?:\.\d{3})+(?:,\d{1,4})?$/.test(abs) || /^\d+,\d{1,4}$/.test(abs)) {
    // Formato AR explícito: 1.015.209,36 / 46612,00 / 4.476.361
    val = parseFloat(abs.replace(/\./g, '').replace(',', '.'));
  } else if (/^\d{1,3}(?:,\d{3})+,\d{2}$/.test(abs)) {
    const idx = abs.lastIndexOf(',');
    val = parseFloat(`${abs.slice(0, idx).replace(/,/g, '')}.${abs.slice(idx + 1)}`);
  } else if (abs.includes(',') && !abs.includes('.')) {
    val = parseFloat(abs.replace(',', '.'));
  } else if (abs.includes('.') && !abs.includes(',')) {
    const parts = abs.split('.');
    val = parts.length > 2 ? parseFloat(abs.replace(/\./g, '')) : parseFloat(abs);
  } else {
    // Ambos separadores: el último determina el decimal
    const lastDot = abs.lastIndexOf('.');
    const lastComma = abs.lastIndexOf(',');
    if (lastComma > lastDot) {
      val = parseFloat(abs.replace(/\./g, '').replace(',', '.'));
    } else {
      val = parseFloat(abs.replace(/,/g, ''));
    }
  }

  if (!Number.isFinite(val)) return 0;
  return round2(isNeg ? -val : val);
};

export const normalizarCuit = (raw: string): string => raw.replace(/\D/g, '');

export const validarCuit = (raw: string): boolean => {
  const digits = normalizarCuit(raw);
  if (!/^\d{11}$/.test(digits)) return false;
  if (/^(\d)\1{10}$/.test(digits)) return false;

  const mult = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2] as const;
  let sum = 0;
  for (let i = 0; i < 10; i++) {
    sum += Number(digits[i]) * (mult[i] ?? 0);
  }
  const rem = 11 - (sum % 11);
  const esperado = rem === 11 ? 0 : rem === 10 ? 9 : rem;
  return Number(digits[10]) === esperado;
};

export const formatearCuit = (digitsRaw: string): string => {
  const d = normalizarCuit(digitsRaw);
  if (d.length !== 11) return digitsRaw.trim();
  return `${d.slice(0, 2)}-${d.slice(2, 10)}-${d.slice(10)}`;
};
