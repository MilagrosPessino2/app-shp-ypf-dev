export type DimensionHW = { alto: number; ancho: number };

export type ImagenOrdenada = {
  data: ArrayBuffer;               // bytes en formato PNG.
  dimension: DimensionHW;          // dimensiones REAJUSTADAS (px)
  dimesionOriginal: { alto: number; ancho: number }; // dimensiones originales (px)
  extension: 'image/png';               // devolvemos PNG 
};

// Firma PNG: 89 50 4E 47 0D 0A 1A 0A
//Comprueba si los primeros 8 bytes corresponden a la firma de PNG.
//Sirve para asegurarse de que lo que generamos realmente es un PNG y no un formato camuflado.
function looksLikePng(u8: Uint8Array): boolean {
  if (u8.byteLength < 8) return false;
  const sig = [0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A];
  for (let i = 0; i < sig.length; i++) if (u8[i] !== sig[i]) return false;
  return true;
}


//Crea una URL temporal con el Blob.
//Carga la imagen en un <img> invisible para obtener su ancho y alto originales.
//Libera la URL temporal después.

async function readNaturalSizeFromBlob(blob: Blob): Promise<{ w: number; h: number }> {
  const url = URL.createObjectURL(blob);
  try {
    return await new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve({ w: img.naturalWidth || 0, h: img.naturalHeight || 0 });
      img.onerror = reject;
      img.src = url;
    });
  } finally {
    URL.revokeObjectURL(url);
  }
}

/* Devuelve un PNG (re-encodado) manteniendo tamaño y con clamp de seguridad */
/*Carga la imagen original y la dibuja en un <canvas>.
Escala la imagen si excede los 8000 px de ancho o alto.
Vuelve a exportarla a un Blob PNG limpio.
Devuelve ese PNG junto con las dimensiones finales. */

async function transcodeToPNG(blob: Blob): Promise<{ out: Blob; w: number; h: number }> {
  let { w, h } = await readNaturalSizeFromBlob(blob);
  if (!w || !h) throw new Error('Dimensiones inválidas');

  const MAX_W = 8000, MAX_H = 8000;
  const scale = Math.min(1, MAX_W / w, MAX_H / h);
  const W = Math.max(1, Math.floor(w * scale));
  const H = Math.max(1, Math.floor(h * scale));

  const canvas = document.createElement('canvas');
  canvas.width = W; canvas.height = H;
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
      canvas.toBlob(b => b ? resolve(b) : reject(new Error('toBlob falló')), 'image/png')
    );
    return { out, w: W, h: H };
  } finally {
    URL.revokeObjectURL(url);
  }
}

/* garantiza PNG valido para Word */
/*Descarga la imagen desde una URL.
La transcodifica siempre a PNG seguro.
Verifica que no esté vacía y que tenga firma PNG.
Devuelve los bytes (ArrayBuffer), alto y ancho ajustados*/ 

export async function loadImageOriginal(
  url: string
): Promise<{ data: ArrayBuffer; alto: number; ancho: number; extension: 'image/png' } | null> {
  const res = await fetch(url);
  if (!res.ok) return null;

  let blob = await res.blob();
  if (!blob.type?.startsWith('image/')) return null;

  let png: Blob, W = 0, H = 0;
  try {
    const tr = await transcodeToPNG(blob);
    png = tr.out; W = tr.w; H = tr.h;
  } catch {
    return null;
  }

  const ab = await png.arrayBuffer();
  const u8 = new Uint8Array(ab);
  if (u8.byteLength === 0 || !looksLikePng(u8)) return null;

  return { data: ab, alto: H, ancho: W, extension: 'image/png' };
}


/* Comparador por alto ASC → ancho ASC */
export function comparaImagenesPorAltoAncho(a: ImagenOrdenada, b: ImagenOrdenada): number {
  const da = a.dimension.alto - b.dimension.alto;
  if (da !== 0) return da;
  return a.dimension.ancho - b.dimension.ancho;
}

/* Inserta manteniendo el orden dado por comparar */
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
