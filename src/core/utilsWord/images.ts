export type DimensionHW = { alto: number; ancho: number };
export type ImagenOrdenada = {
    data: ArrayBuffer; // bytes (PNG seguro)
    dimension: DimensionHW; // dimensiones REAJUSTADAS (px)
    dimensionOriginal: { alto: number; ancho: number };
    extension: 'image/png';
};

/* Verifica si los 8 bytes parecen un PNG (firma) */
function esFormatoPng(u8: Uint8Array): boolean {
    if (u8.byteLength < 8) return false;
    const firmaArchivo = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
    for (let i = 0; i < firmaArchivo.length; i++) 
        if (u8[i] !== firmaArchivo[i]) return false;
    return true;
}

async function obtenerDimensionesOriginalesDeImagen(
    blob: Blob
): Promise<{ w: number; h: number }> {
    const url = URL.createObjectURL(blob);
    try {
        return await new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () =>
                resolve({
                    w: img.naturalWidth || 0,
                    h: img.naturalHeight || 0,
                });
            img.onerror = reject;
            img.src = url;
        });
    } finally {
        URL.revokeObjectURL(url);
    }
}

/* Re-encodifica a PNG (clamp 8000px) */
async function convertirAFormatoPng(
    blob: Blob
): Promise<{ out: Blob; w: number; h: number }> {
    const { w, h } = await obtenerDimensionesOriginalesDeImagen(blob);
    if (!w || !h) throw new Error('Dimensiones inválidas');

    const MAX_W = 8000,
          MAX_H = 8000;
    const scale = Math.min(1, MAX_W / w, MAX_H / h);
    const W = Math.max(1, Math.floor(w * scale));
    const H = Math.max(1, Math.floor(h * scale));

    const canvas = document.createElement('canvas');
    canvas.width = W;
    canvas.height = H;
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('Canvas 2D no disponible');

    const url = URL.createObjectURL(blob);
    try {
        const img = await new Promise<HTMLImageElement>((resolve, reject) => {
            const i = new Image();
            i.onload = () => resolve(i);
            i.onerror = reject;
            i.src = url;
        });
        ctx.drawImage(img, 0, 0, W, H);

        const out = await new Promise<Blob>((resolve, reject) =>
            canvas.toBlob(
                (b) => (b ? resolve(b) : reject(new Error('toBlob falló'))),
                'image/png'
            )
        );
        return { out, w: W, h: H };
    } finally {
        URL.revokeObjectURL(url);
    }
}

/* Descarga + convierte a PNG seguro; devuelve bytes + dimensiones */
export async function descargarYConvertirImagen(url: string): Promise<
    | {
        data: ArrayBuffer;
        alto: number;
        ancho: number;
        extension: 'image/png';
    }
    | undefined
> {
    // Incluir credenciales por si la URL requiere cookies (mismo origen/SharePoint)
    const res = await fetch(url, { credentials: 'include' });
    if (!res.ok) return undefined;

    const blob = await res.blob();
    if (!blob.type?.startsWith('image/')) return undefined;

    let png: Blob,
        W = 0,
        H = 0;
    try {
        const tr = await convertirAFormatoPng(blob);
        png = tr.out;
        W = tr.w;
        H = tr.h;
    } catch {
        return undefined;
    }

    const ab = await png.arrayBuffer();
    const u8 = new Uint8Array(ab);
    if (u8.byteLength === 0 || !esFormatoPng(u8)) return undefined;

    return { data: ab, alto: H, ancho: W, extension: 'image/png' };
}

export function comparaImagenesPorAltoAncho(
    a: ImagenOrdenada,
    b: ImagenOrdenada
): number {
    const da = a.dimension.alto - b.dimension.alto;
    if (da !== 0) return da;
    return a.dimension.ancho - b.dimension.ancho;
}

export function insertarOrdenado<T>(
    coleccion: T[],
    aInsertar: T,
    comparar: (a: T, b: T) => number
): T[] {
    let i = 0;
    while (i < coleccion.length && comparar(aInsertar, coleccion[i]) > 0) i++;
    coleccion.splice(i, 0, aInsertar);
    return coleccion;
}
