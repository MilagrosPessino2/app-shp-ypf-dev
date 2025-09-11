import { Paragraph } from 'docx';
import { htmlToParagraphsControlled } from '../htmlToDoc';



/* Detalle/Descripción del ítem con HTML controlado Paragraph*/
export function noveltyDetail(htmlFragment: string): Paragraph[] {
  return htmlToParagraphsControlled(htmlFragment ?? '');
}
