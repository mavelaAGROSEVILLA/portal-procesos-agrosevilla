const blockElements = new Set([
  'ADDRESS',
  'ARTICLE',
  'BLOCKQUOTE',
  'DIV',
  'H1',
  'H2',
  'H3',
  'H4',
  'H5',
  'H6',
  'LI',
  'OL',
  'P',
  'SECTION',
  'UL'
]);

const getPlainText = (node: Node): string => {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent || '';
  }
  if (node.nodeType !== Node.ELEMENT_NODE) {
    return Array.from(node.childNodes).map(getPlainText).join('');
  }

  const element = node as Element;
  if (element.tagName === 'SCRIPT' || element.tagName === 'STYLE') {
    return '';
  }
  if (element.tagName === 'BR') {
    return '\n';
  }

  const content = Array.from(element.childNodes).map(getPlainText).join('');
  return blockElements.has(element.tagName) ? `\n${content}\n` : content;
};

export const normalizeRichText = (value: string): string => {
  if (!value) {
    return '';
  }

  const parsed = new DOMParser().parseFromString(value, 'text/html');
  return Array.from(parsed.body.childNodes)
    .map(getPlainText)
    .join('')
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
};
