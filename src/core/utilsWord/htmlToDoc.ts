import {
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    AlignmentType,
    WidthType,
    BorderStyle,
} from 'docx';
import { COLOR_TOKENS, TAMANIO } from '../utils/Constants';
import {
    isInlineTag,
    inlineTextContent,
    looksNumericOrMoney,
    readColor,
    stripTags,
} from './htmlUtils';

// Tipo del VALOR del enum AlignmentType (no la clave)
type AlignValue = (typeof AlignmentType)[keyof typeof AlignmentType];

/* -----------------------------------------------------------
 * 1) API COMPATIBLE: se mantiene igual (solo párrafos)
 * ----------------------------------------------------------- */
export function htmlToParagraphsControlled(html: string): Paragraph[] {
    const blocks = htmlToBlocksControlled(html);
    // Filtramos solo Paragraph para no romper a quienes esperan Paragraph[]
    return blocks.filter((b): b is Paragraph => b instanceof Paragraph);
}

/* -----------------------------------------------------------
 * 2) NUEVA API: devuelve Paragraph | Table
 *    (tablas reales y centradas con buen estilo)
 * ----------------------------------------------------------- */
export function htmlToBlocksControlled(html: string): Array<Paragraph | Table> {
    const blocks: Array<Paragraph | Table> = [];
    const safe = (html ?? '').trim();
    if (!safe) return [defaultParagraph('')];

    // Parseo HTML
    let doc: Document | undefined;
    try {
        const parser = new DOMParser();
        doc = parser.parseFromString(safe, 'text/html');
    } catch {
        return [defaultParagraph(stripTags(safe))];
    }

    if (!doc?.body) return [defaultParagraph(stripTags(safe))];
    // Recorremos el body
    for (const node of Array.from(doc.body.childNodes)) {
        const parts = elementToBlocks(node);
        if (parts.length) blocks.push(...parts);
    }
    // Si no quedó nada, devolvemos un párrafo vacío
    return blocks.length ? blocks : [defaultParagraph('')];
}

/* ---------- mapeo a Paragraph | Table (con recursión) ---------- */
function elementToBlocks(node: Node): Array<Paragraph | Table> {
    if (node.nodeType === Node.TEXT_NODE) {
        const text = (node.textContent ?? '').replace(/\s+/g, ' ').trim();
        return text ? [defaultParagraph(text)] : [];
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return [];

    const el = node as HTMLElement;
    const tag = el.tagName.toLowerCase();

    if (tag === 'table') {
        // ahora recibimos (Table | Paragraph)[]
        const tBlocks = buildDocxTableFromHtml(el as HTMLTableElement);
        return tBlocks.length
            ? tBlocks
            : [defaultParagraph(inlineTextContent(el))];
    }

    // Contenedores: bajamos recursivamente y también capturamos su inline como párrafo
    if (
        tag === 'p' ||
        tag === 'div' ||
        tag === 'section' ||
        tag === 'article'
    ) {
        return blocksFromContainer(el);
    }

    if (tag === 'ul' || tag === 'ol') {
        const items = Array.from(el.querySelectorAll(':scope > li'));
        return items.map((li, idx) =>
            paragraphFromInline(li, tag === 'ol' ? `${idx + 1}. ` : '• ')
        );
    }

    if (tag === 'br') return [defaultParagraph('')];

    if (isInlineTag(tag)) return [paragraphFromInline(el)];

    // Otros nodos: tratar como contenedor genérico
    return blocksFromContainer(el);
}

function blocksFromContainer(container: Element): Array<Paragraph | Table> {
    const out: Array<Paragraph | Table> = [];

    // 1) Párrafo del contenido inline del contenedor (si existe)
    const inlineTxt = inlineTextContent(container);
    if (inlineTxt) out.push(paragraphFromInline(container));

    // 2) Recorremos SOLO hijos de tipo "bloque". NO renderizamos hijos inline.
    for (const child of Array.from(container.childNodes)) {
        if (child.nodeType !== Node.ELEMENT_NODE) continue;
        const el = child as HTMLElement;
        const tag = el.tagName.toLowerCase();

        if (tag === 'table') {
            const tBlocks = buildDocxTableFromHtml(el as HTMLTableElement);
            out.push(...tBlocks);
            continue;
        }

        if (tag === 'ul' || tag === 'ol') {
            const items = Array.from(el.querySelectorAll(':scope > li'));
            for (let i = 0; i < items.length; i++) {
                const bullet = tag === 'ol' ? `${i + 1}. ` : '• ';
                out.push(paragraphFromInline(items[i], bullet));
            }
            continue;
        }

        if (
            tag === 'p' ||
            tag === 'div' ||
            tag === 'section' ||
            tag === 'article'
        ) {
            out.push(...blocksFromContainer(el)); // bajar a bloques anidados
            continue;
        }

        // Para cualquier otro elemento que no sea inline ni los de arriba,
        // intentamos bajar (por si contiene bloques dentro)
        if (!isInlineTag(tag)) {
            out.push(...blocksFromContainer(el));
        }
        // Si es inline, lo ignoramos porque ya quedó incluido en el párrafo del contenedor.
    }

    return out;
}

/* ---------- tabla DOCX con estilo, centrada y con ESPACIO DESPUÉS ---------- */
function buildDocxTableFromHtml(
    tableEl: HTMLTableElement
): Array<Table | Paragraph> {
    // Tipado estricto en queries
    const headRows = Array.from(
        tableEl.querySelectorAll<HTMLTableRowElement>(':scope > thead > tr')
    );
    const bodyRows = Array.from(
        tableEl.querySelectorAll<HTMLTableRowElement>(':scope > tbody > tr')
    );
    const directRows = Array.from(
        tableEl.querySelectorAll<HTMLTableRowElement>(':scope > tr')
    );

    // Combinar + deduplicar sin perder el tipo
    const rows: HTMLTableRowElement[] = Array.from(
        new Set<HTMLTableRowElement>([...headRows, ...bodyRows, ...directRows])
    );
    if (!rows.length) return [];

    // Detectar header
    const firstCells = Array.from(
        rows[0].querySelectorAll<HTMLTableCellElement>(
            ':scope > th, :scope > td'
        )
    );
    const firstIsHeader = firstCells.some(
        (c) => c.tagName.toLowerCase() === 'th'
    );

    const docxRows = rows.map((tr, rowIdx) => {
        const cells = Array.from(
            tr.querySelectorAll<HTMLTableCellElement>(
                ':scope > th, :scope > td'
            )
        );
        const isHeaderRow = rowIdx === 0 && firstIsHeader;

        const docxCells = cells.map((c) => {
            const txt = inlineTextContent(c);
            const align: AlignValue | undefined = looksNumericOrMoney(txt)
                ? AlignmentType.RIGHT
                : undefined;

            return new TableCell({
                children: [paragraphFromInline(c, '', { align })],
                margins: { top: 120, bottom: 120, left: 140, right: 140 },
                shading: isHeaderRow ? { fill: 'F2F2F2' } : undefined,
                borders: {
                    top: {
                        style: BorderStyle.SINGLE,
                        size: 2,
                        color: 'D9D9D9',
                    },
                    bottom: {
                        style: BorderStyle.SINGLE,
                        size: 2,
                        color: 'D9D9D9',
                    },
                    left: {
                        style: BorderStyle.SINGLE,
                        size: 2,
                        color: 'D9D9D9',
                    },
                    right: {
                        style: BorderStyle.SINGLE,
                        size: 2,
                        color: 'D9D9D9',
                    },
                },
            });
        });

        return new TableRow({ children: docxCells });
    });

    const table = new Table({
        alignment: AlignmentType.CENTER, // centrada en el documento
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: docxRows,
        borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' },
            left: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' },
            right: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' },
            insideHorizontal: {
                style: BorderStyle.SINGLE,
                size: 2,
                color: 'E6E6E6',
            },
            insideVertical: {
                style: BorderStyle.SINGLE,
                size: 2,
                color: 'E6E6E6',
            },
        },
    });

    // Agregamos un spacer de 6pt (120 twips) DESPUÉS de la tabla
    const spacerAfterTable = new Paragraph({ spacing: { after: 120 } });

    return [table, spacerAfterTable];
}

/* ---------- inline a Paragraph (con alineación opcional) ---------- */
function paragraphFromInline(
    container: Element,
    prefix = '',
    opts?: { align?: AlignValue }
): Paragraph {
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

            if (tag === 'br') {
                runs.push(new TextRun({ text: '', break: 1 }));
                continue;
            }
            if (tag === 'strong' || tag === 'b') {
                runs.push(styledRun(inlineTextContent(el), { bold: true }));
                continue;
            }
            if (tag === 'em' || tag === 'i') {
                runs.push(styledRun(inlineTextContent(el), { italics: true }));
                continue;
            }
            if (tag === 'u') {
                runs.push(styledRun(inlineTextContent(el), { underline: {} }));
                continue;
            }
            if (tag === 'span') {
                const styleColor = readColor(el);
                runs.push(
                    styledRun(inlineTextContent(el), { color: styleColor })
                );
                continue;
            }
            if (tag === 'a') {
                const txt =
                    inlineTextContent(el) || el.getAttribute('href') || '';
                runs.push(styledRun(txt, { underline: {}, color: '0000EE' }));
                continue;
            }

            runs.push(baseRun(inlineTextContent(el)));
        }
    }

    return new Paragraph({
        children: runs.length ? runs : [baseRun('')],
        spacing: { after: 120 },
        alignment: opts?.align, // valor del enum directamente
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
    opts: {
        bold?: boolean;
        italics?: boolean;
        underline?: {};
        color?: string;
    } = {}
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

function defaultParagraph(text: string): Paragraph {
    return new Paragraph({
        children: [baseRun(text)],
        spacing: { after: 120 },
    });
}
