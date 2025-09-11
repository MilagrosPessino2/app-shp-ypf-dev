import { BorderStyle, Paragraph } from 'docx';
import { COLOR_TOKENS, TAMANIO } from '../../utils/Constants';

/** Línea separadora fina entre items */
export function thinSeparator(): Paragraph {
  return new Paragraph({
    border: {
      bottom: { style: BorderStyle.SINGLE, size: TAMANIO.tamanioBorderArea, color: COLOR_TOKENS.separadorLinea },
    },
    spacing: { before: 120, after: 120 },
  });
}
