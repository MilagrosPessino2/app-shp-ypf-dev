// src/core/utilsWord/blocks/noveltyDetail.ts
import { Paragraph, TextRun, TabStopType } from 'docx';
import { htmlToParagraphsControlled } from '../htmlToDoc';

/** Espacio entre párrafos (twips). 240 ≈ 12pt; 120 ≈ 6pt */
const DETAIL_GAP = 240;
/** Posición del tab para listas (twips). 720 = 0.5" */
const LIST_TAB = 720;
/** Detecta líneas con viñeta simple */
const BULLET_RE = /^\s*(?:•|-|\*)\s+/;

/**
 * Convierte texto/HTML de detalle en Paragraphs:
 * - Si es texto plano, cada ENTER crea un párrafo y se agrega espacio entre párrafos.
 * - Las líneas que empiezan con "• ", "-" o "*" se formatean como ítem con tab/indent.
 * - Si es HTML, delega a htmlToParagraphsControlled y agrega espacio entre párrafos intercalando un espaciador.
 */
export function noveltyDetail(htmlOrText: string): Paragraph[] {
    const content = htmlOrText ?? '';
    const hasHtml = /<[a-z][\s\S]*>/i.test(content);

    // Caso A: TEXTO PLANO (multilínea)
    if (!hasHtml) {
        const lines = content.replace(/\r\n/g, '\n').split(/\n+/);
        const out: Paragraph[] = [];

        for (let i = 0; i < lines.length; i++) {
            const raw = lines[i];
            const last = i === lines.length - 1;

            if (BULLET_RE.test(raw)) {
                const text = raw.replace(BULLET_RE, '').trim();
                out.push(
                    new Paragraph({
                        tabStops: [
                            { type: TabStopType.LEFT, position: LIST_TAB },
                        ],
                        indent: { left: LIST_TAB, hanging: LIST_TAB / 2 },
                        spacing: { after: last ? 0 : DETAIL_GAP / 2 },
                        children: [
                            new TextRun('•'),
                            new TextRun('\t'),
                            new TextRun(text),
                        ],
                    })
                );
            } else {
                out.push(
                    new Paragraph({
                        spacing: { after: last ? 0 : DETAIL_GAP },
                        children: [new TextRun(raw)],
                    })
                );
            }
        }
        return out;
    }

    // Caso B: HTML (p, br, b, i, u, ul, li, etc.)
    const paras = htmlToParagraphsControlled(content);

    if (paras.length <= 1) return paras;

    // Intercalar espaciador entre párrafos (sin mutar objetos y sin any)
    const out: Paragraph[] = [];
    for (let i = 0; i < paras.length; i++) {
        out.push(paras[i]);
        if (i < paras.length - 1) {
            out.push(new Paragraph({ spacing: { before: DETAIL_GAP } }));
        }
    }
    return out;
}
