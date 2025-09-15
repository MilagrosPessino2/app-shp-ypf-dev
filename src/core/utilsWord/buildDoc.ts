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
import { htmlToBlocksControlled } from './htmlToDoc';

const PAGE_CONTENT_WIDTH = 500;
const MIN_IMAGE_WIDTH = 0;

/** Párrafo vacío con espacio después (pt → twips) */
function spacer(pt: number): Paragraph {
    return new Paragraph({ spacing: { after: pt * 20 } });
}

function normKey(input?: string): string {
    const s = (input || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const nfd = s.normalize ? s.normalize('NFD') : s;
    return nfd.replace(/[\u0300-\u036f]/g, '');
}
function cmpBase(a: string, b: string): number {
    const A = normKey(a);
    const B = normKey(b);
    return A < B ? -1 : A > B ? 1 : 0;
}

type Novedad = BuildDocInput['novedades'][number];
type WithResumen = { resumenHtml?: string };
function getResumenHtml(n: Novedad): string | undefined {
    return (n as unknown as WithResumen).resumenHtml;
}

export async function createNovedadesDoc(
    input: BuildDocInput
): Promise<Document> {
    const out: (Paragraph | Table)[] = [];

    // 1) Agrupar por Sector
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

    const sectorKeys: string[] = [];
    sectores.forEach((_v, key) => sectorKeys.push(key));
    sectorKeys.sort((ka, kb) =>
        cmpBase(sectores.get(ka)!.label, sectores.get(kb)!.label)
    );

    for (let s = 0; s < sectorKeys.length; s++) {
        const skey = sectorKeys[s];
        const sectorGroup = sectores.get(skey)!;

        out.push(makeareaBox(sectorGroup.label));

        // 2) Agrupar por Área
        const areas = new Map<string, { label: string; items: Novedad[] }>();
        for (const n of sectorGroup.items) {
            const ak = normKey(n.areaNovedad);
            if (!ak) continue;
            if (!areas.has(ak))
                areas.set(ak, {
                    label: (n.areaNovedad || '').trim(),
                    items: [],
                });
            areas.get(ak)!.items.push(n);
        }

        const areaKeys: string[] = [];
        areas.forEach((_v, key) => areaKeys.push(key));
        areaKeys.sort((ka, kb) =>
            cmpBase(areas.get(ka)!.label, areas.get(kb)!.label)
        );

        for (let a = 0; a < areaKeys.length; a++) {
            const akey = areaKeys[a];
            const areaGroup = areas.get(akey)!;

            out.push(areaHeading(areaGroup.label));

            // 3) Ordenar novedades
            const ordenadas = areaGroup.items
                .slice()
                .sort((x, y) => cmpBase(x.tituloNovedad, y.tituloNovedad));

            // 4) Render de cada novedad
            for (let k = 0; k < ordenadas.length; k++) {
                const nov = ordenadas[k];

                out.push(noveltyTitle(nov.tituloNovedad));

                const resumenHtml = getResumenHtml(nov);
                if (typeof resumenHtml === 'string' && resumenHtml.trim()) {
                    // 👇 ahora se usan bloques (Paragraph | Table)
                    out.push(...htmlToBlocksControlled(resumenHtml));
                }

                const detalle = noveltyDetail(nov.detalleNovedad);
                for (const p of detalle) out.push(p);

                // Solo agregamos un espacio pequeño ANTES de las imágenes (6 pt),
                if (nov.imagenesNovedad?.length) {
                    out.push(spacer(5));
                }

                if (nov.imagenesNovedad && nov.imagenesNovedad.length) {
                    const wrappers: ImagenOrdenada[] = [];
                    for (const url of nov.imagenesNovedad) {
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

                            insertarOrdenado(
                                wrappers,
                                {
                                    data: raw.data,
                                    dimension: { alto: height, ancho: width },
                                    dimensionOriginal: {
                                        alto: raw.alto,
                                        ancho: raw.ancho,
                                    },
                                    extension: raw.extension,
                                },
                                comparaImagenesPorAltoAncho
                            );
                        } catch {
                            // ignorar imagen fallida
                        }
                    }

                    if (wrappers.length) {
                        const escaladas: ImgEscalada[] = wrappers.map((it) => ({
                            data: it.data,
                            width: it.dimension.ancho,
                            height: it.dimension.alto,
                            extension: it.extension,
                        }));
                        const gallery = imageGallery(escaladas);
                        for (const g of gallery) out.push(g);
                    }
                }
                // Quitamos el spacer al final de la novedad para no apilar márgenes.
            }

            // Separador solo entre áreas
            const esUltimaArea = a === areaKeys.length - 1;
            if (!esUltimaArea) out.push(thinSeparator());
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
