import JSZip from 'jszip';

export const MAX_XML_UNCOMPRESSED_BYTES = 15 * 1024 * 1024;

export interface EstilosYComentariosXlsx {
  celdasTachadasPorHoja: Map<string, Set<string>>;
  comentariosThreadedPorHoja: Map<string, Map<string, string>>;
}

const cleanInlineText = (s: string): string =>
  s.replace(/[\u200B-\u200D\uFEFF\u00A0]/g, ' ').replace(/\s+/g, ' ').trim();

const decodeXmlText = (s: string): string =>
  s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(parseInt(dec, 10)));

export const limpiarTextoComentarioExcel = (raw: string): string => {
  const normalizedNewlines = raw.replace(/\r\n?/g, '\n').trim();
  if (!normalizedNewlines) return '';

  const esThreaded =
    /\[(?:Threaded comment|Comentario en subproceso)\]/i.test(normalizedNewlines) ||
    /Your spreadsheet software allows you to read/i.test(normalizedNewlines) ||
    /Su software de hoja de c[áa]lculo le permite leer/i.test(normalizedNewlines);

  if (esThreaded) {
    const marker = normalizedNewlines.match(/(?:^|\n)\s*(?:Comment|Comentario)\s*:\s*/i);
    const body =
      marker && marker.index !== undefined
        ? normalizedNewlines.slice(marker.index + marker[0].length)
        : normalizedNewlines;

    const segmentos = body
      .split(/\n\s*(?:Reply|Respuesta)\s*:\s*/i)
      .map((part) =>
        cleanInlineText(
          part
            .replace(/\[(?:Threaded comment|Comentario en subproceso)\]/gi, ' ')
            .replace(/Your spreadsheet software allows you to read[\s\S]*?(?=\n|$)/gi, ' ')
            .replace(/Su software de hoja de c[áa]lculo[\s\S]*?(?=\n|$)/gi, ' ')
            .replace(/(?:Comment|Comentario)\s*:/gi, ' ')
        )
      )
      .filter(Boolean);

    return segmentos.join(' · ');
  }

  return cleanInlineText(normalizedNewlines);
};

const getAttr = (attrs: string, name: string): string | undefined => {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const m = attrs.match(new RegExp(`(?:^|\\s)${escaped}="([^"]*)"`, 'i'));
  return m?.[1];
};

const colLettersToIndex = (letters: string): number => {
  let idx = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    idx = idx * 26 + (upper.charCodeAt(i) - 64);
  }
  return idx - 1;
};

const hasActiveStrike = (xml: string): boolean => {
  for (const m of xml.matchAll(/<strike\b([^>]*)\/?>/gi)) {
    const val = getAttr(m[1] ?? '', 'val');
    if (val === undefined || (val !== '0' && val.toLowerCase() !== 'false')) {
      return true;
    }
  }
  return false;
};

export const leerArchivoZipTextoSeguro = async (
  zip: JSZip,
  internalPath: string,
  maxBytes: number = MAX_XML_UNCOMPRESSED_BYTES
): Promise<string | null> => {
  const cleanTarget = internalPath.replace(/^\/+/, '').replace(/\\/g, '/');
  const entry =
    zip.file(cleanTarget) ??
    Object.values(zip.files).find(
      (f) => !f.dir && f.name.replace(/^\/+/, '').replace(/\\/g, '/').toLowerCase() === cleanTarget.toLowerCase()
    );
  if (!entry) return null;

  const bytes = await entry.async('uint8array');
  if (bytes.byteLength > maxBytes) {
    throw new Error('El archivo contiene una estructura interna demasiado grande.');
  }
  return new TextDecoder('utf-8').decode(bytes);
};

const parseStruckFontIndices = (stylesXml: string): Set<number> => {
  const struck = new Set<number>();
  const fontsBlock = stylesXml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1];
  if (!fontsBlock) return struck;

  let idx = 0;
  for (const m of fontsBlock.matchAll(/<font\b[^>]*\/>|<font\b[^>]*>[\s\S]*?<\/font>/gi)) {
    if (hasActiveStrike(m[0])) {
      struck.add(idx);
    }
    idx++;
  }
  return struck;
};

const parseStruckXfIndices = (stylesXml: string, struckFonts: ReadonlySet<number>): Set<number> => {
  const struck = new Set<number>();
  if (struckFonts.size === 0) return struck;
  const xfsBlock = stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1];
  if (!xfsBlock) return struck;

  let idx = 0;
  for (const m of xfsBlock.matchAll(/<xf\b([^>]*)(?:\/>|>[\s\S]*?<\/xf>)/gi)) {
    const attrs = m[1] ?? '';
    const fontIdRaw = getAttr(attrs, 'fontId');
    const applyFont = getAttr(attrs, 'applyFont');
    if (fontIdRaw !== undefined && applyFont !== '0') {
      const fontId = parseInt(fontIdRaw, 10);
      if (struckFonts.has(fontId)) {
        struck.add(idx);
      }
    }
    idx++;
  }
  return struck;
};

const parseStruckSharedStrings = (ssXml: string): Set<number> => {
  const struck = new Set<number>();
  let idx = 0;
  for (const m of ssXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)) {
    const siBody = m[1] ?? '';
    const runs = Array.from(siBody.matchAll(/<r\b[^>]*>([\s\S]*?)<\/r>/gi));
    if (runs.length > 0) {
      let runsConTexto = 0;
      let runsTachados = 0;
      for (const r of runs) {
        const rBody = r[1] ?? '';
        const text = cleanInlineText(
          Array.from(rBody.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi))
            .map((tm) => decodeXmlText(tm[1] ?? ''))
            .join('')
        );
        if (!text) continue;
        runsConTexto++;
        const rPr = rBody.match(/<rPr\b[^>]*>([\s\S]*?)<\/rPr>/i)?.[1] ?? '';
        if (hasActiveStrike(rPr)) {
          runsTachados++;
        }
      }
      if (runsConTexto > 0 && runsConTexto === runsTachados) {
        struck.add(idx);
      }
    }
    idx++;
  }
  return struck;
};

const resolveRelativeZipPath = (baseDir: string, target: string): string => {
  const normalizedTarget = target.replace(/\\/g, '/');
  if (normalizedTarget.startsWith('/')) return normalizedTarget.slice(1);
  const parts = baseDir ? baseDir.split('/').filter(Boolean) : [];
  for (const seg of normalizedTarget.split('/')) {
    if (!seg || seg === '.') continue;
    if (seg === '..') parts.pop();
    else parts.push(seg);
  }
  return parts.join('/');
};

export const extraerEstilosYComentariosXlsx = async (
  bufferOrBytes: ArrayBuffer | Uint8Array
): Promise<EstilosYComentariosXlsx> => {
  const empty: EstilosYComentariosXlsx = {
    celdasTachadasPorHoja: new Map(),
    comentariosThreadedPorHoja: new Map()
  };

  const bytes = bufferOrBytes instanceof Uint8Array ? bufferOrBytes : new Uint8Array(bufferOrBytes);
  if (
    bytes.byteLength < 4 ||
    bytes[0] !== 0x50 ||
    bytes[1] !== 0x4b ||
    bytes[2] !== 0x03 ||
    bytes[3] !== 0x04
  ) {
    return empty;
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return empty;
  }
  const hasWorkbook =
    Boolean(zip.file('xl/workbook.xml')) ||
    Object.keys(zip.files).some((k) => k.toLowerCase() === 'xl/workbook.xml');
  if (!hasWorkbook) return empty;

  const [stylesXml, ssXml, workbookXml, wbRelsXml] = await Promise.all([
    leerArchivoZipTextoSeguro(zip, 'xl/styles.xml'),
    leerArchivoZipTextoSeguro(zip, 'xl/sharedStrings.xml'),
    leerArchivoZipTextoSeguro(zip, 'xl/workbook.xml'),
    leerArchivoZipTextoSeguro(zip, 'xl/_rels/workbook.xml.rels')
  ]);

  const struckFonts = stylesXml ? parseStruckFontIndices(stylesXml) : new Set<number>();
  const struckXfs = stylesXml ? parseStruckXfIndices(stylesXml, struckFonts) : new Set<number>();
  const struckSharedStrings = ssXml ? parseStruckSharedStrings(ssXml) : new Set<number>();

  const relTargets = new Map<string, string>();
  if (wbRelsXml) {
    for (const m of wbRelsXml.matchAll(/<Relationship\b([^>]*)\/?>/gi)) {
      const attrs = m[1] ?? '';
      const id = getAttr(attrs, 'Id');
      const target = getAttr(attrs, 'Target');
      if (id && target) {
        relTargets.set(id, resolveRelativeZipPath('xl', target));
      }
    }
  }

  const sheetEntries: Array<{ name: string; path: string }> = [];
  if (workbookXml) {
    let fallbackIdx = 1;
    for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/?>/gi)) {
      const attrs = m[1] ?? '';
      const nameRaw = getAttr(attrs, 'name');
      const rId = getAttr(attrs, 'r:id') ?? getAttr(attrs, 'id');
      if (!nameRaw) continue;
      const name = decodeXmlText(nameRaw);
      const sheetPath = (rId ? relTargets.get(rId) : undefined) ?? `xl/worksheets/sheet${fallbackIdx}.xml`;
      sheetEntries.push({ name, path: sheetPath });
      fallbackIdx++;
    }
  }

  const celdasTachadasPorHoja = new Map<string, Set<string>>();
  const comentariosThreadedPorHoja = new Map<string, Map<string, string>>();

  for (const sheet of sheetEntries) {
    const sheetXml = await leerArchivoZipTextoSeguro(zip, sheet.path);
    if (!sheetXml) continue;

    if (struckXfs.size > 0 || struckSharedStrings.size > 0 || /<strike\b/i.test(sheetXml)) {
      const tachadas = new Set<string>();
      for (const cm of sheetXml.matchAll(/<c\b([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/gi)) {
        const attrs = cm[1] ?? '';
        const body = cm[2] ?? '';
        const refAttr = getAttr(attrs, 'r');
        if (!refAttr) continue;
        const refMatch = refAttr.match(/^([A-Z]+)(\d+)$/i);
        if (!refMatch || !refMatch[1] || !refMatch[2]) continue;

        const row0 = parseInt(refMatch[2], 10) - 1;
        const col0 = colLettersToIndex(refMatch[1]);
        const sAttr = getAttr(attrs, 's');
        const tAttr = getAttr(attrs, 't');

        let esTachada = false;
        if (sAttr !== undefined && struckXfs.has(parseInt(sAttr, 10))) {
          esTachada = true;
        } else if (tAttr === 's' && body) {
          const vMatch = body.match(/<v\b[^>]*>(\d+)<\/v>/i);
          if (vMatch && vMatch[1] && struckSharedStrings.has(parseInt(vMatch[1], 10))) {
            esTachada = true;
          }
        } else if (tAttr === 'inlineStr' && body && hasActiveStrike(body)) {
          const runs = Array.from(body.matchAll(/<r\b[^>]*>([\s\S]*?)<\/r>/gi));
          if (runs.length === 0) {
            esTachada = true;
          } else {
            const conTexto = runs.filter((r) => cleanInlineText(r[1] ?? '').length > 0);
            if (conTexto.length > 0 && conTexto.every((r) => hasActiveStrike(r[1] ?? ''))) {
              esTachada = true;
            }
          }
        }

        if (esTachada) {
          tachadas.add(`${row0}:${col0}`);
        }
      }
      if (tachadas.size > 0) {
        celdasTachadasPorHoja.set(sheet.name, tachadas);
      }
    }

    // Buscar threadedComments vinculados desde las rels de esta hoja
    const slashIdx = sheet.path.lastIndexOf('/');
    const dir = slashIdx >= 0 ? sheet.path.slice(0, slashIdx) : '';
    const base = slashIdx >= 0 ? sheet.path.slice(slashIdx + 1) : sheet.path;
    const sheetRelsPath = dir ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`;
    const sheetRelsXml = await leerArchivoZipTextoSeguro(zip, sheetRelsPath);
    if (sheetRelsXml) {
      const threadedTargets: string[] = [];
      for (const rm of sheetRelsXml.matchAll(/<Relationship\b([^>]*)\/?>/gi)) {
        const target = getAttr(rm[1] ?? '', 'Target');
        if (target && /threadedcomment/i.test(target)) {
          threadedTargets.push(resolveRelativeZipPath(dir, target));
        }
      }

      if (threadedTargets.length > 0) {
        const mapaHoja = new Map<string, string>();
        for (const tcPath of threadedTargets) {
          const tcXml = await leerArchivoZipTextoSeguro(zip, tcPath);
          if (!tcXml) continue;
          for (const tm of tcXml.matchAll(/<threadedComment\b([^>]*)>([\s\S]*?)<\/threadedComment>/gi)) {
            const ref = getAttr(tm[1] ?? '', 'ref');
            const textRaw = tm[2]?.match(/<text\b[^>]*>([\s\S]*?)<\/text>/i)?.[1] ?? '';
            if (!ref || !textRaw) continue;
            const refMatch = ref.match(/^([A-Z]+)(\d+)$/i);
            if (!refMatch || !refMatch[1] || !refMatch[2]) continue;
            const coord = `${parseInt(refMatch[2], 10) - 1}:${colLettersToIndex(refMatch[1])}`;
            const limpio = limpiarTextoComentarioExcel(decodeXmlText(textRaw));
            if (!limpio) continue;
            const prev = mapaHoja.get(coord);
            mapaHoja.set(coord, prev && !prev.includes(limpio) ? `${prev} · ${limpio}` : limpio);
          }
        }
        if (mapaHoja.size > 0) {
          comentariosThreadedPorHoja.set(sheet.name, mapaHoja);
        }
      }
    }
  }

  return {
    celdasTachadasPorHoja,
    comentariosThreadedPorHoja
  };
};

export const extraerCeldasTachadasXlsx = async (
  bufferOrBytes: ArrayBuffer | Uint8Array
): Promise<Map<string, Set<string>>> => {
  const { celdasTachadasPorHoja } = await extraerEstilosYComentariosXlsx(bufferOrBytes);
  return celdasTachadasPorHoja;
};
