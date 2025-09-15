import { Document, Packer } from 'docx';
import { saveAs } from 'file-saver';
import { DOCX_FILENAME } from '../utils/Constants';


/**
 * Empaqueta un Document a .docx y lo descarga automáticamente.
 * Usa file-saver y NO abre el blob en otra pestaña (evita doble descarga).
 */
export async function downloadDoc(
    doc: Document,
    filename = DOCX_FILENAME
): Promise<void> {
    const blob = await Packer.toBlob(doc);
    saveAs(blob, filename);
}
