/**
 * XML Parser utilities using fast-xml-parser
 */

import { XMLParser } from 'fast-xml-parser';

// Type for parsed XML node (JSON from fast-xml-parser)
type XmlNode = {
  [key: string]: any;
  '@_attributes'?: Record<string, string>;
  '#text'?: string;
};

// DOM-like wrapper for JSON nodes
export interface XmlElement {
  localName: string;
  originalTagName?: string; // Original tag name with namespace prefix (e.g., "w:tblBorders")
  textContent: string | null;
  attributes: Map<string, string>;
  childNodes: XmlElement[];
  nodeType: number;
  parent?: XmlElement;
  _data: XmlNode;
}

export class XmlParser {
  private parser: XMLParser;

  constructor() {
    this.parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      textNodeName: '#text',
      preserveOrder: true,  // MUST use true to preserve document element order
      parseTagValue: false,
      trimValues: false,
      removeNSPrefix: false,
      isArray: () => false
    });
  }

  parseFromString(xmlString: string): XmlElement {
    const json = this.parser.parse(xmlString);

    // With preserveOrder: true, the result is an array
    if (Array.isArray(json)) {
      // Find the root element (skip XML declaration if present)
      const rootIndex = json.findIndex(item => {
        const keys = Object.keys(item);
        return keys.length > 0 && keys.some(k => !k.startsWith('?xml') && k !== ':@');
      });

      if (rootIndex >= 0) {
        const rootItem = json[rootIndex];
        // Find the tag name (not :@)
        const rootKey = Object.keys(rootItem).find(k => k !== ':@');
        if (rootKey) {
          return this.jsonToElementFromOrderedArray(rootItem[rootKey], rootKey, rootItem[':@']);
        }
      }
    }

    // Fallback
    return this.jsonToElement({}, 'root');
  }

  private jsonToElementFromOrderedArray(jsonArray: any, tagName: string = 'root', attributes?: any): XmlElement {
    // Handle the ordered array format from preserveOrder: true
    // Format: element's children are an array of items, each item is either:
    //   - {#text: "..."} for text nodes
    //   - {tagName: [...], ":@": {...}} for child elements
    const element: XmlElement = {
      localName: this.getLocalName(tagName),
      originalTagName: tagName.includes(':') ? tagName : undefined,
      textContent: null,
      attributes: new Map(),
      childNodes: [],
      nodeType: 1,
      _data: {}
    };

    // Process attributes
    if (attributes) {
      for (const [key, value] of Object.entries(attributes)) {
        const attrName = key.startsWith('@_') ? key.substring(2) : key;
        element.attributes.set(attrName, String(value));
      }
    }

    if (!Array.isArray(jsonArray)) {
      // Not an array, might be a simple value
      if (typeof jsonArray === 'string' && jsonArray.length > 0) {
        element.textContent = jsonArray; // Preserve whitespace-only strings
      }
      return element;
    }

    // Process array items in order
    for (const item of jsonArray) {
      if (typeof item === 'object' && item !== null) {
        // Check for direct text node  
        if (item['#text'] !== undefined) {
          const text = String(item['#text']);
          // Preserve ALL text content including whitespace-only text!
          // Spaces are critical for proper word spacing in Word documents
          // Only set text if this element doesn't have child elements yet
          if (text.length > 0 && element.childNodes.length === 0) {
            element.textContent = text;
          }
          continue;
        }

        // Process child elements
        for (const [key, value] of Object.entries(item)) {
          if (key === ':@') {
            // Skip attributes marker
            continue;
          }

          // This is a child element
          const childAttrs = item[':@'];
          const child = this.jsonToElementFromOrderedArray(value, key, childAttrs);
          child.parent = element;
          element.childNodes.push(child);
        }
      }
    }

    return element;
  }

  private jsonToElement(json: any, tagName: string = 'root'): XmlElement {
    const element: XmlElement = {
      localName: this.getLocalName(tagName),
      originalTagName: tagName.includes(':') ? tagName : undefined,
      textContent: null,
      attributes: new Map(),
      childNodes: [],
      nodeType: 1,
      _data: {}
    };

    // Handle empty elements: fast-xml-parser returns "" for <tag/>
    if (typeof json === 'string') {
      // If it's an empty string, check if it should be treated as an empty element
      // Empty elements like <w:b/> return "" but should be element nodes
      if (json === '' || json.trim() === '') {
        // For empty strings with element-like tagNames (containing colon or known elements),
        // treat as empty element node, not text node
        if (tagName.includes(':') || tagName.length > 0) {
          element._data = {};
          // Continue processing as element (will have empty attributes and childNodes)
        } else {
          // Unknown empty string, treat as empty text
          element.textContent = '';
          element.nodeType = 3;
          return element;
        }
      } else {
        // Non-empty string is a text node
        element.textContent = json;
        element.nodeType = 3; // Text node
        return element;
      }
    }

    if (json && typeof json === 'object') {
      // Use json directly as data
      const data: XmlNode = json;

      element._data = data;

      // Extract attributes
      // fast-xml-parser stores attributes with @_ prefix plus full attribute name
      // e.g., @_w:val, @_w:ascii, etc.
      for (const [key, value] of Object.entries(data)) {
        if (key.startsWith('@_')) {
          // Remove @_ prefix to get the attribute name
          const attrName = key.substring(2); // Remove '@_'
          element.attributes.set(attrName, String(value));
        }
      }

      // Also check for @_attributes object (fallback, though fast-xml-parser doesn't use this)
      if (data['@_attributes']) {
        for (const [key, value] of Object.entries(data['@_attributes'])) {
          element.attributes.set(key, String(value));
        }
      }

      // Extract text content
      if (data['#text'] !== undefined) {
        element.textContent = String(data['#text']);
      }

      // Extract child elements
      for (const [key, value] of Object.entries(data)) {
        // Skip attributes (keys starting with @_), text content, and special keys
        if (key.startsWith('@_') || key === '#text') continue;

        if (Array.isArray(value)) {
          for (const item of value) {
            // Handle arrays of strings (e.g., multiple <w:t> elements)
            if (typeof item === 'string') {
              const child: XmlElement = {
                localName: this.getLocalName(key),
                originalTagName: key.includes(':') ? key : undefined,
                textContent: item,
                attributes: new Map(),
                childNodes: [],
                nodeType: 1,
                _data: {},
                parent: element
              };
              element.childNodes.push(child);
            } else {
              const child = this.jsonToElement(item, key);
              child.parent = element;
              element.childNodes.push(child);
            }
          }
        } else if (value && typeof value === 'object') {
          const child = this.jsonToElement(value, key);
          child.parent = element;
          element.childNodes.push(child);
        } else if (typeof value === 'string') {
          // Handle text-only elements: fast-xml-parser returns string for <w:t>text</w:t>
          // If key looks like an element name (contains colon), create element with text content
          if (key.includes(':')) {
            const child: XmlElement = {
              localName: this.getLocalName(key),
              originalTagName: key,
              textContent: value,
              attributes: new Map(),
              childNodes: [],
              nodeType: 1, // Element node, but with text content
              _data: {}
            };
            // If the value is empty, it's an empty element like <w:b/>
            if (value === '') {
              child.textContent = null;
            }
            child.parent = element;
            element.childNodes.push(child);
          }
        }
      }
    }

    return element;
  }

  private getLocalName(tagName: string): string {
    // Remove namespace prefix (e.g., "w:body" -> "body")
    const colonIndex = tagName.indexOf(':');
    return colonIndex >= 0 ? tagName.substring(colonIndex + 1) : tagName;
  }

  elements(elem: XmlElement | any, localName?: string): XmlElement[] {
    const result: XmlElement[] = [];

    if (!elem) return result;

    if (elem.nodeType === 9) { // Document node (if ever used)
      // For document, use documentElement
      if (elem.documentElement) {
        elem = elem.documentElement as XmlElement;
      } else {
        return result;
      }
    }

    const children = elem.childNodes || [];
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (child.nodeType === 1) { // Element node
        if (!localName || child.localName === localName) {
          result.push(child);
        }
      }
    }

    return result;
  }

  element(elem: XmlElement | any, localName: string): XmlElement | null {
    const elements = this.elements(elem, localName);
    return elements.length > 0 ? elements[0] : null;
  }

  attr(elem: XmlElement, localName: string, namespace?: string): string | null {
    if (!elem.attributes) return null;

    // Try direct attribute lookup (local name only)
    let attr = elem.attributes.get(localName);
    if (attr !== undefined) return attr;

    // Try with namespace prefix (e.g., "w:val", "r:embed")
    // fast-xml-parser preserves namespace prefixes in attribute names
    const commonPrefixes = ['w', 'r', 'a', 'pic', 'wp'];
    for (const prefix of commonPrefixes) {
      attr = elem.attributes.get(`${prefix}:${localName}`);
      if (attr !== undefined) return attr;
      // Also try without colon separator (some formats)
      attr = elem.attributes.get(`${prefix}${localName}`);
      if (attr !== undefined) return attr;
    }

    // Try case-insensitive and various formats
    for (const [key, value] of elem.attributes.entries()) {
      const keyLower = key.toLowerCase();
      const localNameLower = localName.toLowerCase();

      // Exact match after removing namespace
      if (keyLower === localNameLower) return value as string;

      // Match with any prefix
      if (keyLower.endsWith(`:${localNameLower}`) || keyLower.endsWith(`${localNameLower}`)) {
        const lastColon = keyLower.lastIndexOf(':');
        if (lastColon >= 0 && keyLower.substring(lastColon + 1) === localNameLower) {
          return value as string;
        }
      }
    }

    // Try common namespaces for this attribute
    const commonNamespaces: Record<string, string[]> = {
      'val': [
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'http://schemas.openxmlformats.org/drawingml/2006/main'
      ],
      'w': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'type': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'sz': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'name': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'color': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'fill': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'styleId': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'default': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'before': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'after': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'line': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'lineRule': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'left': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'right': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'firstLine': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'hanging': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'ascii': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'hAnsi': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'eastAsia': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'cs': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'asciiTheme': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'eastAsiaTheme': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'hAnsiTheme': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'cstheme': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'themeColor': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'themeShade': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'themeTint': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'typeface': ['http://schemas.openxmlformats.org/drawingml/2006/main'],
      'abstractNumId': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'numId': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'ilvl': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'hint': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'pos': ['http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
      'embed': [
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'http://schemas.openxmlformats.org/package/2006/relationships'
      ],
      'link': ['http://schemas.openxmlformats.org/officeDocument/2006/relationships']
    };

    if (commonNamespaces[localName]) {
      // Try with namespace prefixes
      for (const prefix of ['w', 'r', 'a', 'pic', 'wp']) {
        const prefixedAttr = `${prefix}:${localName}`;
        const value = elem.attributes.get(prefixedAttr);
        if (value !== undefined) return value;
      }
    }

    return null;
  }

  boolAttr(elem: XmlElement, localName: string, defaultValue = false): boolean {
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

  intAttr(elem: XmlElement, localName: string, defaultValue = 0): number {
    const val = this.attr(elem, localName);
    return val ? parseInt(val, 10) : defaultValue;
  }

  lengthAttr(elem: XmlElement, localName: string, unit: 'dxa' | 'pt' | 'emu' = 'dxa'): number | null {
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

  // Helper to serialize element back to XML string (replaces XMLSerializer)
  serializeToString(elem: XmlElement): string {
    return this.elementToXml(elem);
  }

  private elementToXml(elem: XmlElement, indent = 0): string {
    if (elem.nodeType === 3) { // Text node
      return elem.textContent || '';
    }

    let xml = '';
    const indentStr = '  '.repeat(indent);
    // Use original tag name with prefix if available, otherwise map from localName
    const tagName = elem.originalTagName || this.getFullTagName(elem.localName);

    // Build attributes
    let attrs = '';
    if (elem.attributes) {
      for (const [key, value] of elem.attributes.entries()) {
        attrs += ` ${key}="${this.escapeXml(value)}"`;
      }
    }

    // Handle self-closing or empty elements
    if (elem.childNodes.length === 0 && !elem.textContent) {
      return `${indentStr}<${tagName}${attrs}/>`;
    }

    xml += `${indentStr}<${tagName}${attrs}>`;

    if (elem.textContent) {
      xml += this.escapeXml(elem.textContent);
    }

    for (const child of elem.childNodes) {
      if (child.nodeType === 1) {
        xml += '\n' + this.elementToXml(child, indent + 1);
      } else if (child.nodeType === 3) {
        xml += this.escapeXml(child.textContent || '');
      }
    }

    if (elem.childNodes.length > 0) {
      xml += '\n' + indentStr;
    }

    xml += `</${tagName}>`;

    return xml;
  }

  private getFullTagName(localName: string): string {
    // Common namespace prefixes for DOCX
    const namespaceMap: Record<string, string> = {
      'body': 'w:body',
      'p': 'w:p',
      'r': 'w:r',
      't': 'w:t',
      'pPr': 'w:pPr',
      'rPr': 'w:rPr',
      'tbl': 'w:tbl',
      'tr': 'w:tr',
      'tc': 'w:tc',
      'tblPr': 'w:tblPr',
      'b': 'w:b',
      'i': 'w:i',
      'u': 'w:u',
      'strike': 'w:strike',
      'sz': 'w:sz',
      'rFonts': 'w:rFonts',
      'color': 'w:color',
      'highlight': 'w:highlight',
      'vertAlign': 'w:vertAlign',
      'spacing': 'w:spacing',
      'jc': 'w:jc',
      'ind': 'w:ind',
      'pStyle': 'w:pStyle',
      'numPr': 'w:numPr',
      'numId': 'w:numId',
      'ilvl': 'w:ilvl',
      'tab': 'w:tab',
      'br': 'w:br',
      'drawing': 'w:drawing',
      'inline': 'wp:inline',
      'graphic': 'a:graphic',
      'graphicData': 'a:graphicData',
      'pic': 'pic:pic',
      'blipFill': 'pic:blipFill',
      'blip': 'a:blip',
      'extent': 'wp:extent',
      'tcPr': 'w:tcPr',
      'gridSpan': 'w:gridSpan',
      'vMerge': 'w:vMerge',
      'tcW': 'w:tcW',
      'shd': 'w:shd',
      'vAlign': 'w:vAlign',
      'tcBorders': 'w:tcBorders',
      'tcMar': 'w:tcMar',
      'tblGrid': 'w:tblGrid',
      'gridCol': 'w:gridCol',
      'tblBorders': 'w:tblBorders',
      'tblInd': 'w:tblInd',
      'top': 'w:top',
      'left': 'w:left',
      'bottom': 'w:bottom',
      'right': 'w:right',
      'insideH': 'w:insideH',
      'insideV': 'w:insideV',
      'tblStyleRowBandSize': 'w:tblStyleRowBandSize',
      'tblStyleColBandSize': 'w:tblStyleColBandSize',
      'tblW': 'w:tblW',
      'tblLook': 'w:tblLook',
      'tblStyle': 'w:tblStyle'
    };

    return namespaceMap[localName] || localName;
  }

  private escapeXml(str: string): string {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}