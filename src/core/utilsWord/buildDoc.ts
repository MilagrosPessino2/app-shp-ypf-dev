import { Document, Paragraph, Table } from 'docx';
import {
    buildFooter,
    buildHeader,
    makeareaBox,
    areaHeading,
    noveltyDetail,
    noveltyTitle,
    thinSeparator,
} from './blocks';
import { imageGallery, type ImgEscalada } from './blocks/imageGallery';
import {
    loadImageOriginal,
    insertarOrdenado,
    comparaImagenesPorAltoAncho,
    type ImagenOrdenada,
} from './images';
import type { BuildDocInput } from './types';

const PAGE_CONTENT_WIDTH = 500;
const MIN_IMAGE_WIDTH = 0;

export async function createNovedadesDoc(
    input: BuildDocInput
): Promise<Document> {
    const out: (Paragraph | Table)[] = []; // ✅ Acepta Paragraphs y Tables (como makeareaBox)

    // 1. Agrupar por sectorGeneral
    const sectorMap = new Map<string, BuildDocInput['novedades']>();

    for (const nov of input.novedades) {
        if (!sectorMap.has(nov.sectorGeneral)) {
            sectorMap.set(nov.sectorGeneral, []);
        }
        sectorMap.get(nov.sectorGeneral)!.push(nov);
    }

    // 2. Procesar por grupo
    const sectoresOrdenados = Array.from(sectorMap.entries()).sort(([a], [b]) =>
        a.localeCompare(b)
    );

    for (const [sector, novedades] of sectoresOrdenados) {
        // Mostrar el area box UNA VEZ
        out.push(makeareaBox(sector));

        // Ordenar novedades por areaNovedad, luego por titulo
        const ordenadas = [...novedades].sort((a, b) => {
            const areaCmp = a.areaNovedad.localeCompare(b.areaNovedad);
            if (areaCmp !== 0) return areaCmp;
            return a.tituloNovedad.localeCompare(b.tituloNovedad);
        });

        for (const nov of ordenadas) {
            // Título del área
            out.push(areaHeading(nov.areaNovedad));

            // Título de la novedad
            out.push(noveltyTitle(nov.tituloNovedad));

            // Detalle enriquecido
            const detalleParas = noveltyDetail(nov.detalleNovedad);
            out.push(...detalleParas);

            // Imágenes
            if (nov.imagenesNovedad && nov.imagenesNovedad.length > 0) {
                const wrappersOrdenados: ImagenOrdenada[] = [];

                for (const url of nov.imagenesNovedad) {
                    try {
                        const raw = await loadImageOriginal(url);
                        if (!raw) continue;

                        const relacion =
                            raw.ancho > 0 ? raw.alto / raw.ancho : 0;
                        if (!(relacion > 0 && isFinite(relacion))) continue;

                        const naturalMax = Math.min(
                            Math.max(1, raw.ancho),
                            Math.max(1, PAGE_CONTENT_WIDTH)
                        );

                        const width =
                            MIN_IMAGE_WIDTH > 0
                                ? Math.min(
                                      naturalMax,
                                      Math.max(1, MIN_IMAGE_WIDTH)
                                  )
                                : naturalMax;

                        const height = Math.max(
                            1,
                            Math.round(width * relacion)
                        );

                        const wrapper: ImagenOrdenada = {
                            data: raw.data,
                            dimension: { alto: height, ancho: width },
                            dimensionOriginal: {
                                alto: raw.alto,
                                ancho: raw.ancho,
                            },
                            extension: raw.extension,
                        };

                        insertarOrdenado(
                            wrappersOrdenados,
                            wrapper,
                            comparaImagenesPorAltoAncho
                        );
                    } catch {
                        // ignorar imagen fallida
                    }
                }

                if (wrappersOrdenados.length > 0) {
                    const escaladas: ImgEscalada[] = wrappersOrdenados.map(
                        (w) => ({
                            data: w.data,
                            width: w.dimension.ancho,
                            height: w.dimension.alto,
                            extension: w.extension,
                        })
                    );

                    const galleryParas = imageGallery(escaladas);
                    out.push(...galleryParas);
                }
            }

            // Separador entre novedades
            out.push(thinSeparator());
            out.push(new Paragraph({ spacing: { after: 50 } }));
        }
    }

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
                headers: {
                    default: buildHeader(input.confidentialityLabel ?? ''),
                },
                footers: {
                    default: buildFooter(input.confidentialityLabel ?? ''),
                },
                children: out,
            },
        ],
    });
}
