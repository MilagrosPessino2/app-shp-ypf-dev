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

/** Normaliza: case-insensitive + accent-insensitive + trim + colapsa espacios */
function normKey(input?: string): string {
    const s = (input || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const nfd = s.normalize ? s.normalize('NFD') : s;
    return nfd.replace(/[\u0300-\u036f]/g, '');
}

/** Comparación usando la misma normalización */
function cmpBase(a: string, b: string): number {
    const A = normKey(a);
    const B = normKey(b);
    return A < B ? -1 : A > B ? 1 : 0;
}

type Novedad = BuildDocInput['novedades'][number];
type WithResumen = { resumenHtml?: string };

export async function createNovedadesDoc(
    input: BuildDocInput
): Promise<Document> {
    const out: (Paragraph | Table)[] = [];

    // 1) Agrupar por Sector General (CI + AI)
    const sectores = new Map<string, { label: string; items: Novedad[] }>();
    for (let i = 0; i < input.novedades.length; i++) {
        const n = input.novedades[i];
        const k = normKey(n.sectorGeneral);
        if (!k) continue;
        if (!sectores.has(k))
            sectores.set(k, {
                label: (n.sectorGeneral || '').trim(),
                items: [],
            });
        sectores.get(k)!.items.push(n);
    }

    // Ordenar sectores por label normalizado
    const sectorKeys = Array.from(sectores.keys()).sort((ka, kb) =>
        cmpBase(sectores.get(ka)!.label, sectores.get(kb)!.label)
    );

    for (let s = 0; s < sectorKeys.length; s++) {
        const skey = sectorKeys[s];
        const sectorGroup = sectores.get(skey)!;

        // Contenedor del sector (mostrar la primera etiqueta encontrada)
        out.push(makeareaBox(sectorGroup.label));

        // 2) Agrupar por Área (CI + AI)
        const areas = new Map<string, { label: string; items: Novedad[] }>();
        const novedadesSector = sectorGroup.items;

        for (let j = 0; j < novedadesSector.length; j++) {
            const n = novedadesSector[j];
            const ak = normKey(n.areaNovedad);
            if (!ak) continue;
            if (!areas.has(ak))
                areas.set(ak, {
                    label: (n.areaNovedad || '').trim(),
                    items: [],
                });
            areas.get(ak)!.items.push(n);
        }

        // Ordenar áreas por label normalizado
        const areaKeys = Array.from(areas.keys()).sort((ka, kb) =>
            cmpBase(areas.get(ka)!.label, areas.get(kb)!.label)
        );

        for (let a = 0; a < areaKeys.length; a++) {
            const akey = areaKeys[a];
            const areaGroup = areas.get(akey)!;

            // Título del Área UNA vez
            out.push(areaHeading(areaGroup.label));

            // 3) Ordenar novedades dentro del área (por título, CI + AI)
            const ordenadas = areaGroup.items
                .slice()
                .sort((x, y) => cmpBase(x.tituloNovedad, y.tituloNovedad));

            // 4) Render de cada novedad
            for (let k = 0; k < ordenadas.length; k++) {
                const nov = ordenadas[k];

                // Título
                out.push(noveltyTitle(nov.tituloNovedad));

                // Resumen HTML (si viene)
                const resumenHtml = (nov as unknown as WithResumen).resumenHtml;
                if (typeof resumenHtml === 'string' && resumenHtml.trim()) {
                    const resumenParas =
                        htmlToParagraphsControlled(resumenHtml);
                    out.push(...resumenParas);
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
                                PAGE_CONTENT_WIDTH
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

            // 5) Separador SOLO entre Áreas
            const esUltimaArea = a === areaKeys.length - 1;
            if (!esUltimaArea) {
                out.push(thinSeparator());
            }
        }
    }

    return new Document({
        styles: {
            default: {
                document: {
                    run: { font: 'Calibri' },
                    // ✅ En docx@9.5.1 el spacing por defecto va dentro de "document"
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
