/**
 * Tipo mínimo requerido para poder calcular área.
 */
export type Dimension = {
    width: number;
    height: number;
};

/**
 * Comparator: ordena por área ASC.
 * - Si el área es igual, ordena por alto ASC.
 * - Si también es igual, ordena por ancho ASC.
 */
export function compareByArea(a: Dimension, b: Dimension): number {
    const areaA = a.width * a.height;
    const areaB = b.width * b.height;

    if (areaA !== areaB) return areaA - areaB;
    if (a.height !== b.height) return a.height - b.height;
    return a.width - b.width;
}
