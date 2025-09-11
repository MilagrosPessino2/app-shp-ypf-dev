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
  sectorGeneral: string;
  novedad: SeccionArea[];
  maxItemsPerSection?: number;
  confidentialityLabel?: string;
}


