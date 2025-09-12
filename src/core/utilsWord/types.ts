/** Item dentro de un área de novedad */
export interface ItemNovedad {
    tituloNovedad: string;
    detalleNovedad: string;
    imagenesNovedad?: string[];
}
export interface SeccionArea {
    areaNovedad: string;
    items: ItemNovedad[];
}

/** Entrada para construir el documento de novedades */
export interface BuildDocInput {
    novedades: {
        sectorGeneral: string;
        areaNovedad: string;
        tituloNovedad: string;
        resumenHtml: string;
        detalleNovedad: string;
        imagenesNovedad?: string[];
    }[];
    confidentialityLabel?: string;
}
export type NovedadItem = {
    sectorGeneral: string;
    tituloNovedad: string;
    resumenHtml: string;
    detalleNovedad: string;
    imagenesNovedad: string[];
};
