/**
 * XML Parser utilities (simplified from docxjs)
 */

export class XmlParser {
  elements(elem: Element | Document, localName?: string): Element[] {
    const result: Element[] = [];

    const children = elem.childNodes;
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (child.nodeType === 1) { // Element node
        const el = child as Element;
        if (!localName || el.localName === localName) {
          result.push(el);
        }
      }
    }

    return result;
  }

  element(elem: Element, localName: string): Element | null {
    const elements = this.elements(elem, localName);
    return elements.length > 0 ? elements[0] : null;
  }

  attr(elem: Element, localName: string, namespace?: string): string | null {
    // Try without namespace first
    let attr = elem.attributes.getNamedItem(localName);
    if (attr) return attr.value;

    // Try with null namespace
    attr = elem.attributes.getNamedItemNS(null, localName);
    if (attr) return attr.value;

    // Try with specified namespace
    if (namespace) {
      attr = elem.attributes.getNamedItemNS(namespace, localName);
      if (attr) return attr.value;
    }

    // Try common namespaces for this attribute
    const commonNamespaces: Record<string, string[]> = {
      'val': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main', // w:val
        'http://schemas.openxmlformats.org/drawingml/2006/main' // a:val
      ],
      'w': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:w (width)
      ],
      'type': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:type
      ],
      'sz': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:sz (size)
      ],
      'name': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:name
      ],
      'color': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:color
      ],
      'fill': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:fill
      ],
      'styleId': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:styleId
      ],
      'default': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:default
      ],
      'before': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:before (spacing)
      ],
      'after': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:after (spacing)
      ],
      'line': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:line (spacing)
      ],
      'lineRule': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:lineRule
      ],
      'left': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:left (indentation)
      ],
      'right': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:right (indentation)
      ],
      'firstLine': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:firstLine
      ],
      'hanging': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:hanging
      ],
      'ascii': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:ascii (font)
      ],
      'hAnsi': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:hAnsi (font)
      ],
      'eastAsia': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:eastAsia (font)
      ],
      'cs': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:cs (complex script font)
      ],
      'asciiTheme': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:asciiTheme (theme font reference)
      ],
      'eastAsiaTheme': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:eastAsiaTheme
      ],
      'hAnsiTheme': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:hAnsiTheme
      ],
      'cstheme': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:cstheme
      ],
      'themeColor': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:themeColor
      ],
      'themeShade': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:themeShade
      ],
      'themeTint': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:themeTint
      ],
      'typeface': [
        'http://schemas.openxmlformats.org/drawingml/2006/main' // a:typeface (font name in theme)
      ],
      'abstractNumId': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:abstractNumId
      ],
      'numId': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:numId
      ],
      'ilvl': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:ilvl (numbering level)
      ],
      'hint': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:hint (font hint)
      ],
      'pos': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' // w:pos (tab position)
      ],
      'embed': [
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'http://schemas.openxmlformats.org/package/2006/relationships'
      ],
      'link': [
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
      ]
    };

    if (commonNamespaces[localName]) {
      for (const ns of commonNamespaces[localName]) {
        attr = elem.attributes.getNamedItemNS(ns, localName);
        if (attr) return attr.value;
      }
    }

    return null;
  }

  boolAttr(elem: Element, localName: string, defaultValue = false): boolean {
    const val = this.attr(elem, localName);
    if (val === null) return defaultValue;

    switch (val) {
      case '1':
      case 'true':
      case 'on':
        return true;
      case '0':
      case 'false':
      case 'off':
        return false;
      default:
        return defaultValue;
    }
  }

  intAttr(elem: Element, localName: string, defaultValue = 0): number {
    const val = this.attr(elem, localName);
    return val ? parseInt(val, 10) : defaultValue;
  }

  lengthAttr(elem: Element, localName: string, unit: 'dxa' | 'pt' | 'emu' = 'dxa'): number | null {
    const val = this.attr(elem, localName);
    if (!val) return null;

    const num = parseInt(val, 10);

    // Convert to points
    switch (unit) {
      case 'dxa': // twips (1/1440 inch)
        return num / 20; // 20 twips = 1 point
      case 'emu': // EMUs
        return num / 12700;
      case 'pt':
        return num;
      default:
        return num / 20;
    }
  }
}

