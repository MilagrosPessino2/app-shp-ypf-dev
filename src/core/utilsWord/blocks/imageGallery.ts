// src/core/utilsWord/blocks/imageGallery.ts
import { AlignmentType, ImageRun, Paragraph } from 'docx';

export type ImgEscalada = {
    data: ArrayBuffer;
    width: number;
    height: number;
    extension?: 'image/png' | 'image/jpeg';
};

// Tipo = unión de los valores del enum AlignmentType
type AlignmentValue = (typeof AlignmentType)[keyof typeof AlignmentType];

export type GalleryOptions = {
    /** Espacio antes de la PRIMER imagen (twips). Ej: 240 ≈ 12pt */
    topGapTwips?: number;
    /** Espacio después de CADA imagen (twips). Default: 200 */
    afterEach?: number;
    /** Alineación (default: center) */
    align?: AlignmentValue;
};

export function imageGallery(
    images: ImgEscalada[],
    opts: GalleryOptions = {}
): Paragraph[] {
    const topGap = Math.max(0, opts.topGapTwips ?? 0);
    const after = Math.max(0, opts.afterEach ?? 200);
    const align: AlignmentValue = opts.align ?? AlignmentType.CENTER;

    return images.map(({ data, width, height, extension }, idx) => {
        const w = Math.max(1, Math.round(width));
        const h = Math.max(1, Math.round(height));
        const docxType: 'png' | 'jpg' =
            extension === 'image/jpeg' ? 'jpg' : 'png';

        return new Paragraph({
            alignment: align,
            spacing: idx === 0 ? { before: topGap, after } : { after },
            children: [
                new ImageRun({
                    data: new Uint8Array(data),
                    type: docxType,
                    transformation: { width: w, height: h },
                }),
            ],
        });
    });
}
