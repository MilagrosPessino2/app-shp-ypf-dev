import { Paragraph, TextRun } from 'docx';
import { COLOR_TOKENS, TAMANIO } from '../utils/Constants';


export function htmlToParagraphsControlled(html: string): Paragraph[] {
  const paragraphs: Paragraph[] = [];
  const safe = (html ?? '').trim();
  if (!safe) return [defaultParagraph('')];

  let doc: Document | null = null;
  try {
    const parser = new DOMParser();
    doc = parser.parseFromString(safe, 'text/html');
  } catch {
    return [defaultParagraph(stripTags(safe))];
  }
  if (!doc || !doc.body) return [defaultParagraph(stripTags(safe))];

  for (const node of Array.from(doc.body.childNodes)) {
    const ps = elementToParagraphs(node);
    if (ps.length) paragraphs.push(...ps);
  }
  return paragraphs.length ? paragraphs : [defaultParagraph('')];
}

/* ---------- mapeo ---------- */
function elementToParagraphs(node: Node): Paragraph[] {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = (node.textContent ?? '').replace(/\s+/g, ' ').trim();
    return text ? [defaultParagraph(text)] : [];
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return [];

  const el = node as HTMLElement;
  const tag = el.tagName.toLowerCase();

  if (tag === 'p' || tag === 'div') return [paragraphFromInline(el)];
  if (tag === 'br') return [defaultParagraph('')];

  if (isInlineTag(tag)) return [paragraphFromInline(el)];

  if (tag === 'ul' || tag === 'ol') {
    const items = Array.from(el.querySelectorAll(':scope > li'));
    return items.map((li, idx) => {
      const bullet = tag === 'ol' ? `${idx + 1}. ` : '• ';
      return paragraphFromInline(li, bullet);
    });
  }

  if (tag === 'table') {
    const rows = Array.from(el.querySelectorAll(':scope > tbody > tr, :scope > tr'));
    if (!rows.length) return [paragraphFromInline(el)];
    return rows.map((tr) => {
      const cells = Array.from(tr.querySelectorAll(':scope > td, :scope > th'));
      const text = cells.map((c) => inlineTextContent(c)).join(' | ');
      return defaultParagraph(text);
    });
  }

  if (tag === 'img') return [defaultParagraph('')]; // las imágenes las manejás fuera

  return [defaultParagraph(inlineTextContent(el))];
}

function paragraphFromInline(container: Element, prefix = ''): Paragraph {
  const runs: TextRun[] = [];
  if (prefix) runs.push(baseRun(prefix));

  for (const child of Array.from(container.childNodes)) {
    if (child.nodeType === Node.TEXT_NODE) {
      const t = (child.textContent ?? '').replace(/\s+/g, ' ');
      if (t) runs.push(baseRun(t));
      continue;
    }
    if (child.nodeType === Node.ELEMENT_NODE) {
      const el = child as HTMLElement;
      const tag = el.tagName.toLowerCase();

      if (tag === 'br') { runs.push(new TextRun({ text: '', break: 1 })); continue; }
      if (tag === 'strong' || tag === 'b') { runs.push(styledRun(inlineTextContent(el), { bold: true })); continue; }
      if (tag === 'em' || tag === 'i') { runs.push(styledRun(inlineTextContent(el), { italics: true })); continue; }
      if (tag === 'u') { runs.push(styledRun(inlineTextContent(el), { underline: {} })); continue; }
      if (tag === 'span') {
        const styleColor = readColor(el);
        runs.push(styledRun(inlineTextContent(el), { color: styleColor }));
        continue;
      }
      if (tag === 'a') {
        const txt = inlineTextContent(el) || el.getAttribute('href') || '';
        runs.push(styledRun(txt, { underline: {}, color: '0000EE' }));
        continue;
      }

      runs.push(baseRun(inlineTextContent(el)));
    }
  }

  return new Paragraph({
    children: runs.length ? runs : [baseRun('')],
    spacing: { after: 120 },
  });
}

/* ---------- estilos corporativos ---------- */
function baseRun(text: string): TextRun {
  return new TextRun({
    text,
    color: COLOR_TOKENS.textoDetalleNovedad,
    size: TAMANIO.tamanioTextoDetalleNovedad,
  });
}

function styledRun(
  text: string,
  opts: { bold?: boolean; italics?: boolean; underline?: {}; color?: string } = {}
): TextRun {
  return new TextRun({
    text,
    bold: opts.bold,
    italics: opts.italics,
    underline: opts.underline,
    color: opts.color ?? COLOR_TOKENS.textoDetalleNovedad,
    size: TAMANIO.tamanioTextoDetalleNovedad,
  });
}

/* ---------- utilitarios ---------- */
function isInlineTag(tag: string): boolean {
  return ['span', 'strong', 'b', 'em', 'i', 'u', 'a', 'br'].includes(tag);
}

function inlineTextContent(el: Element): string {
  const walker = el.ownerDocument!.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  const parts: string[] = [];
  let n: Node | null;
  // eslint-disable-next-line no-cond-assign
  while ((n = walker.nextNode())) {
    const t = (n.textContent ?? '').replace(/\s+/g, ' ');
    if (t) parts.push(t);
  }
  return parts.join('').trim();
}

function readColor(el: HTMLElement): string | undefined {
  const style = el.getAttribute('style') || '';
  const m = /color\s*:\s*(#[0-9a-fA-F]{3,8})/i.exec(style);
  return m ? m[1].replace('#', '').toUpperCase() : undefined;
}

function stripTags(s: string): string {
  return s.replace(/<[^>]*>/g, '');
}

function defaultParagraph(text: string): Paragraph {
  return new Paragraph({
    children: [baseRun(text)],
    spacing: { after: 120 },
  });
}
