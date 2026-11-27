import JSZip from 'jszip';

export type FormatoArchivoDetectado =
  | 'XLSX'
  | 'DOCX'
  | 'XLS'
  | 'PDF'
  | 'JPEG'
  | 'PNG'
  | 'WEBP'
  | 'CSV'
  | 'DESCONOCIDO';

const startsWithBytes = (bytes: Uint8Array, prefix: ReadonlyArray<number>, offset = 0): boolean => {
  if (bytes.length < offset + prefix.length) return false;
  for (let i = 0; i < prefix.length; i++) {
    if (bytes[offset + i] !== prefix[i]) return false;
  }
  return true;
};

const esTextoSinNul = (bytes: Uint8Array): boolean => {
  const limite = Math.min(bytes.length, 4096);
  let controlCount = 0;
  for (let i = 0; i < limite; i++) {
    const b = bytes[i] ?? 0;
    if (b === 0) return false;
    if (b < 9 || (b > 13 && b < 32)) {
      controlCount++;
    }
  }
  return limite > 0 && controlCount <= Math.floor(limite * 0.02);
};

export const detectarFormatoPorFirma = async (
  bytes: Uint8Array,
  nombre: string
): Promise<FormatoArchivoDetectado> => {
  if (bytes.length === 0) return 'DESCONOCIDO';

  // %PDF-
  if (startsWithBytes(bytes, [0x25, 0x50, 0x44, 0x46, 0x2d])) {
    return 'PDF';
  }
  // JPEG: FF D8 FF
  if (startsWithBytes(bytes, [0xff, 0xd8, 0xff])) {
    return 'JPEG';
  }
  // PNG: 89 50 4E 47 0D 0A 1A 0A
  if (startsWithBytes(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    return 'PNG';
  }
  // WEBP: RIFF....WEBP
  if (
    startsWithBytes(bytes, [0x52, 0x49, 0x46, 0x46], 0) &&
    startsWithBytes(bytes, [0x57, 0x45, 0x42, 0x50], 8)
  ) {
    return 'WEBP';
  }
  // OLE Compound Document (.xls / .doc heredado): D0 CF 11 E0 A1 B1 1A E1
  if (startsWithBytes(bytes, [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])) {
    return 'XLS';
  }
  // OOXML (ZIP PK\x03\x04 o PK\x05\x06)
  if (
    startsWithBytes(bytes, [0x50, 0x4b, 0x03, 0x04]) ||
    startsWithBytes(bytes, [0x50, 0x4b, 0x05, 0x06])
  ) {
    try {
      const zip = await JSZip.loadAsync(bytes);
      if (zip.file('word/document.xml')) return 'DOCX';
      if (zip.file('xl/workbook.xml')) return 'XLSX';
      // Algunos XLSX minimalistas pueden usar uppercase o casing mixto
      const names = Object.keys(zip.files).map((k) => k.toLowerCase());
      if (names.some((k) => k === 'word/document.xml')) return 'DOCX';
      if (names.some((k) => k === 'xl/workbook.xml')) return 'XLSX';
      return 'DESCONOCIDO';
    } catch {
      return 'DESCONOCIDO';
    }
  }

  const ext = nombre.split('.').pop()?.toLowerCase() ?? '';
  if ((ext === 'csv' || ext === 'txt') && esTextoSinNul(bytes)) {
    return 'CSV';
  }

  return 'DESCONOCIDO';
};

const FORMATO_TO_ETQ: Record<FormatoArchivoDetectado, string> = {
  XLSX: 'un Excel (.xlsx)',
  DOCX: 'un documento Word (.docx)',
  XLS: 'un Excel (.xls)',
  PDF: 'un PDF',
  JPEG: 'una imagen',
  PNG: 'una imagen',
  WEBP: 'una imagen',
  CSV: 'una planilla CSV',
  DESCONOCIDO: 'un archivo desconocido'
};

export const extensionCoincide = (
  nombre: string,
  formatoReal: FormatoArchivoDetectado
): { coincide: boolean; aviso?: string } => {
  const ext = nombre.split('.').pop()?.toLowerCase() ?? '';
  if (!ext || formatoReal === 'DESCONOCIDO') return { coincide: true };

  const esImg = formatoReal === 'JPEG' || formatoReal === 'PNG' || formatoReal === 'WEBP';
  const extImg = ext === 'jpg' || ext === 'jpeg' || ext === 'png' || ext === 'webp';

  const coincide =
    (formatoReal === 'XLSX' && (ext === 'xlsx' || ext === 'xlsm')) ||
    (formatoReal === 'XLS' && (ext === 'xls' || ext === 'xlsx')) ||
    (formatoReal === 'DOCX' && ext === 'docx') ||
    (formatoReal === 'PDF' && ext === 'pdf') ||
    (formatoReal === 'CSV' && (ext === 'csv' || ext === 'txt')) ||
    (esImg && extImg);

  if (coincide) return { coincide: true };
  return {
    coincide: false,
    aviso: `El archivo tenía extensión .${ext} pero su contenido real es ${FORMATO_TO_ETQ[formatoReal]}.`
  };
};
