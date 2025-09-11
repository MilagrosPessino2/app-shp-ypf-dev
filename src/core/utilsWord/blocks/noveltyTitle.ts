import { Paragraph, TextRun, UnderlineType } from 'docx';
import { COLOR_TOKENS, TAMANIO } from '../../utils/Constants';

/** Título de novedad (ítem) */
export function noveltyTitle(text: string): Paragraph {
    return new Paragraph({
        children: [
            new TextRun({
                text: text.toUpperCase(),
                bold: true,
                color: COLOR_TOKENS.tituloItemNovedad,
                size: TAMANIO.tamanioTittle,
                underline: { type: UnderlineType.NONE },
            }),
        ],
        spacing: { after: 60 },
    });
}
