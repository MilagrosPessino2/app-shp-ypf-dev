// src/core/utilsWord/htmlUtils.ts

/**
 * Retorna si un tag HTML es considerado inline para renderizado.
 */
export function isInlineTag(tag: string): boolean {
    return ['span', 'strong', 'b', 'em', 'i', 'u', 'a', 'br'].includes(tag);
}

/**
 * Extrae texto plano de los nodos de texto de un elemento HTML.
 */
export function inlineTextContent(el: Element): string {
    const walker = el.ownerDocument!.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    const parts: string[] = [];
    let n: Node | null;
    while ((n = walker.nextNode())) {
        const t = (n.textContent ?? '').replace(/\s+/g, ' ');
        if (t) parts.push(t);
    }
    return parts.join('').trim();
}

/**
 * Detecta si un string representa un número o valor monetario.
 */
export function looksNumericOrMoney(text: string): boolean {
    const s = text
        .replace(/\u2212/g, '-')
        .replace(/\s/g, '')
        .trim();
    const money =
        /^((US\$)|(\$)|(€)|(£))?-?\d{1,3}([.,']\d{3})*([.,]\d+)?$|^-?\d+([.,]\d+)?$/i;
    return money.test(s);
}

/**
 * Extrae el color hexadecimal (sin `#`) desde el atributo `style` de un elemento.
 */
export function readColor(el: HTMLElement): string | undefined {
    const style = el.getAttribute('style') || '';
    const m = /color\s*:\s*(#[0-9a-fA-F]{3,8})/i.exec(style);
    return m ? m[1].replace('#', '').toUpperCase() : undefined;
}

/**
 * Elimina todas las etiquetas HTML de un string.
 */
export function stripTags(s: string): string {
    return s.replace(/<[^>]*>/g, '');
}
