import { AlignmentType, ImageRun, Paragraph } from 'docx';

export type ImgEscalada = {
  data: ArrayBuffer;                          // bytes de imagen
  width: number;                              // ancho a renderizar (px)
  height: number;                             // alto a renderizar (px)
  extension?: 'image/png' | 'image/jpeg';     // tipo
};

export function imageGallery(images: ImgEscalada[]): Paragraph[] {
  return images.map(({ data, width, height, extension }) => {
    const w = Math.max(1, Math.round(width));
    const h = Math.max(1, Math.round(height));

    // `docx` espera 'jpg' (no 'jpeg')
    const docxType: 'png' | 'jpg' = extension === 'image/jpeg' ? 'jpg' : 'png';

    return new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [
        new ImageRun({
          data: new Uint8Array(data),
          type: docxType,                      // <-- 'png' | 'jpg'
          transformation: { width: w, height: h },
        }),
      ],
    });
  });
}
