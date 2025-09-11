import { Document, Paragraph } from 'docx';
import {
  buildFooter,
  buildHeader,
  makeareaBox,
  areaHeading,
  noveltyDetail,   // devuelve Paragraph[]
  noveltyTitle,
  thinSeparator,
} from './blocks';
import { imageGallery, type ImgEscalada } from './blocks/imageGallery';
import type { SeccionArea, BuildDocInput } from './types';
import {
  loadImageOriginal,
  insertarOrdenado,
  comparaImagenesPorAltoAncho,
  type ImagenOrdenada,
} from './images';

async function buildSectionsAsync(
  sections: SeccionArea[],
  pageContentWidthPx: number,
  minImageWidthPx = 0
): Promise<Paragraph[]> {
  const out: Paragraph[] = [];

  for (const { areaNovedad, items } of sections) {
    out.push(areaHeading(areaNovedad));

    for (const { tituloNovedad, detalleNovedad, imagenesNovedad } of items) {
      out.push(noveltyTitle(tituloNovedad));

      // Detalle con HTML controlado 
      const detalleParas = noveltyDetail(detalleNovedad);
      out.push(...detalleParas);

      //  Imágenes 
      if (imagenesNovedad && imagenesNovedad.length > 0) {
        const wrappersOrdenados: ImagenOrdenada[] = [];

        for (const url of imagenesNovedad) {
          try {
            const raw = await loadImageOriginal(url);
            if (!raw) continue;

            const relacion = raw.ancho > 0 ? raw.alto / raw.ancho : 0;
            if (!(relacion > 0 && isFinite(relacion))) continue;

            const naturalMax = Math.min(
              Math.max(1, raw.ancho),
              Math.max(1, pageContentWidthPx)
            );

            const width =
              minImageWidthPx > 0
                ? Math.min(naturalMax, Math.max(1, minImageWidthPx))
                : naturalMax;

            const height = Math.max(1, Math.round(width * relacion));

            const wrapper: ImagenOrdenada = {
              data: raw.data,
              dimension: { alto: height, ancho: width },
              dimesionOriginal: { alto: raw.alto, ancho: raw.ancho },
              extension: raw.extension,
            };

            insertarOrdenado(wrappersOrdenados, wrapper, comparaImagenesPorAltoAncho);
          } catch {
            /* ignorar imagen fallida */
          }
        }

        if (wrappersOrdenados.length > 0) {
          const escaladas: ImgEscalada[] = wrappersOrdenados.map((w) => ({
            data: w.data,
            width: w.dimension.ancho,
            height: w.dimension.alto,
            extension: w.extension,
          }));

          const galleryParas = imageGallery(escaladas);
         
          out.push(...galleryParas);
        }
      }

      out.push(thinSeparator());
    }

    out.push(new Paragraph({ spacing: { after: 50 } }));
  }

  return out;
}

/* Builder principal (async) */
export async function createNovedadesDoc(input: BuildDocInput): Promise<Document> {
  const { sectorGeneral, novedad } = input;

  const PAGE_CONTENT_WIDTH = 500; // ajustá si cambian márgenes
  const MIN_IMAGE_WIDTH = 0;

  const sectionChildren = [
    makeareaBox(sectorGeneral),
    ...(await buildSectionsAsync(novedad, PAGE_CONTENT_WIDTH, MIN_IMAGE_WIDTH)),
  ];

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: 'Calibri' },
          paragraph: { spacing: { before: 80, after: 80 } },
        },
      },
    },
    sections: [
      {
        headers: { default: buildHeader('YPF-Confidencial') },
        footers: { default: buildFooter('YPF-Confidencial') },
        children: sectionChildren,
      },
    ],
  });
}

export default createNovedadesDoc;
