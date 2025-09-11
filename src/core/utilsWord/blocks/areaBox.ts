import {
    AlignmentType,
    BorderStyle,
    Paragraph,
    ShadingType,
    Table,
    TableCell,
    TableLayoutType,
    TableRow,
    TextRun,
    WidthType,
} from 'docx';
import { COLOR_TOKENS, TAMANIO } from '../../utils/Constants';


/** Franja "area" con fondo y borde del sector */
export function makeareaBox(sectorGeneral: string): Table {
    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        alignment: AlignmentType.LEFT,
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.SINGLE,
                                size: TAMANIO.tamanioBorderArea,
                                color: COLOR_TOKENS.bordeArea,
                            },
                            bottom: {
                                style: BorderStyle.SINGLE,
                                size: TAMANIO.tamanioBorderArea,
                                color: COLOR_TOKENS.bordeArea,
                            },
                            left: {
                                style: BorderStyle.SINGLE,
                                size: TAMANIO.tamanioBorderArea,
                                color: COLOR_TOKENS.bordeArea,
                            },
                            right: {
                                style: BorderStyle.SINGLE,
                                size: TAMANIO.tamanioBorderArea,
                                color: COLOR_TOKENS.bordeArea,
                            },
                        },
                        shading: {
                            type: ShadingType.CLEAR,
                            fill: COLOR_TOKENS.fondoArea,
                            color: 'auto',
                        },
                        margins: { top: 60, bottom: 60, left: 80, right: 80 },
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: sectorGeneral,
                                        bold: true,
                                        color: COLOR_TOKENS.textoNegro,
                                        size: TAMANIO.tamanioTextoArea,
                                    }),
                                ],
                                spacing: { before: 0, after: 0 },
                            }),
                        ],
                    }),
                ],
            }),
        ],
    });
}
