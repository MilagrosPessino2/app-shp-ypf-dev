import { AlignmentType, Footer, Paragraph, TextRun } from 'docx';
import { COLOR_TOKENS, TAMANIO } from '../../utils/Constants';

/** Footer reutilizable (centrado) */
export function buildFooter(label: string): Footer {
    return new Footer({
        children: [
            new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                    new TextRun({
                        text: label,
                        color: COLOR_TOKENS.textoDetalleNovedad,
                        size: TAMANIO.tamanioTextoFooterHeader,
                    }),
                ],
            }),
        ],
    });
}
