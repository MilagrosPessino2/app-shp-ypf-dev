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
import { htmlToParagraphsControlled } from './htmlToDoc';

const PAGE_CONTENT_WIDTH = 500;
const MIN_IMAGE_WIDTH = 0;

export async function createNovedadesDoc(
    input: BuildDocInput
): Promise<Document> {
    const out: (Paragraph | Table)[] = [];

    // 1) Agrupar por Sector General
    const sectores = new Map<string, BuildDocInput['novedades']>();
    for (let i = 0; i < input.novedades.length; i++) {
        const n = input.novedades[i];
        if (!sectores.has(n.sectorGeneral)) sectores.set(n.sectorGeneral, []);
        sectores.get(n.sectorGeneral)!.push(n);
    }

    // Ordenar sectores por nombre SIN iteradores: usar keys()
    const sectoresKeys = Array.from(sectores.keys()).sort((a, b) =>
        a.localeCompare(b)
    );

    for (let s = 0; s < sectoresKeys.length; s++) {
        const sector = sectoresKeys[s];
        const novedadesSector = sectores.get(sector)!;

        // Mostrar el contenedor del sector UNA sola vez
        out.push(makeareaBox(sector));

        // 2) Dentro del sector, agrupar por Área de novedad
        const areas = new Map<string, BuildDocInput['novedades']>();
        for (let j = 0; j < novedadesSector.length; j++) {
            const n = novedadesSector[j];
            if (!areas.has(n.areaNovedad)) areas.set(n.areaNovedad, []);
            areas.get(n.areaNovedad)!.push(n);
        }

        // Ordenar áreas alfabéticamente SIN iteradores
        const areasKeys = Array.from(areas.keys()).sort((a, b) =>
            a.localeCompare(b)
        );

        for (let a = 0; a < areasKeys.length; a++) {
            const area = areasKeys[a];
            const novedadesArea = areas.get(area)!;

            // Mostrar el título del Área UNA vez
            out.push(areaHeading(area));

            // Ordenar novedades dentro del área (por título)
            const ordenadas = novedadesArea
                .slice()
                .sort((x, y) => x.tituloNovedad.localeCompare(y.tituloNovedad));

            // 3) Render de cada novedad dentro del área (sin separador entre novedades)
            for (let k = 0; k < ordenadas.length; k++) {
                const nov = ordenadas[k];

                // Título de la novedad
                out.push(noveltyTitle(nov.tituloNovedad));
                
                // resumen HTML
                if (nov.resumenHtml && nov.resumenHtml.trim()) {
                const resumenParas = htmlToParagraphsControlled(nov.resumenHtml);
                out.push(...resumenParas, new Paragraph({ spacing: { after: 80 } })); 
                }

                // Detalle enriquecido
                const detalle = noveltyDetail(nov.detalleNovedad);
                for (let d = 0; d < detalle.length; d++) out.push(detalle[d]);

                // Imágenes (si hay)
                if (nov.imagenesNovedad && nov.imagenesNovedad.length) {
                    const wrappers: ImagenOrdenada[] = [];
                    for (let u = 0; u < nov.imagenesNovedad.length; u++) {
                        const url = nov.imagenesNovedad[u];
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
                                wrappers,
                                wrapper,
                                comparaImagenesPorAltoAncho
                            );
                        } catch {
                            // ignorar imagen fallida
                        }
                    }

                    if (wrappers.length) {
                        const escaladas: ImgEscalada[] = new Array(
                            wrappers.length
                        );
                        for (let w = 0; w < wrappers.length; w++) {
                            const it = wrappers[w];
                            escaladas[w] = {
                                data: it.data,
                                width: it.dimension.ancho,
                                height: it.dimension.alto,
                                extension: it.extension,
                            };
                        }
                        const gallery = imageGallery(escaladas);
                        for (let g = 0; g < gallery.length; g++)
                            out.push(gallery[g]);
                    }
                }
            }

            // 4) Separador SOLO entre Áreas (no entre novedades)
            const esUltimaArea = a === areasKeys.length - 1;
            if (!esUltimaArea) {
                out.push(thinSeparator());
                // out.push(new Paragraph({ spacing: { after: 50 } }));
            }
        }
    }

    return new Document({
        styles: {
            default: {
                document: {
                    run: { font: 'Calibri' },
                    paragraph: { spacing: { before: 0, after: 0 } },
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
