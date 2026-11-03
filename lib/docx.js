/**
 * HTML/text -> DOCX generation and HTML -> plain text, without a browser DOM.
 *
 * `htmlToDocx` supports the block and inline subset a document converter
 * usually meets: headings, paragraphs, lists, blockquotes, code blocks,
 * tables, bold/italic/code/links. Images are intentionally out of scope for
 * this version and are skipped with their alt text kept.
 */

import {
  AlignmentType,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType
} from 'docx'
import { parse } from 'node-html-parser'

const HEADINGS = {
  h1: HeadingLevel.HEADING_1,
  h2: HeadingLevel.HEADING_2,
  h3: HeadingLevel.HEADING_3,
  h4: HeadingLevel.HEADING_4,
  h5: HeadingLevel.HEADING_5,
  h6: HeadingLevel.HEADING_6
}

const BLOCK_TAGS = new Set([
  'address', 'article', 'aside', 'blockquote', 'div', 'dl', 'fieldset',
  'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3', 'h4', 'h5',
  'h6', 'header', 'hr', 'li', 'main', 'nav', 'ol', 'p', 'pre', 'section',
  'table', 'tbody', 'td', 'tfoot', 'th', 'thead', 'tr', 'ul'
])

function isElement(node) {
  return node && node.nodeType === 1
}

function isWhitespaceText(node) {
  return node?.nodeType === 3 && String(node.textContent ?? node.rawText ?? '').trim() === ''
}

function tagOf(node) {
  return isElement(node) ? String(node.rawTagName || '').toLowerCase() : ''
}

function inlineRuns(node, style = {}) {
  if (!node) return []
  if (node.nodeType === 3) {
    const text = node.textContent ?? node.rawText ?? ''
    if (!text) return []
    return [new TextRun({ text, ...style })]
  }
  if (!isElement(node)) return []

  const tag = tagOf(node)
  if (tag === 'br') return [new TextRun({ break: 1 })]
  if (tag === 'script' || tag === 'style' || tag === 'head') return []

  if (tag === 'img') {
    const alt = node.getAttribute('alt')
    return alt ? [new TextRun({ text: alt, italics: true })] : []
  }

  const next = { ...style }
  if (tag === 'strong' || tag === 'b') next.bold = true
  if (tag === 'em' || tag === 'i') next.italics = true
  if (tag === 'u') next.underline = {}
  if (tag === 's' || tag === 'del' || tag === 'strike') next.strike = true
  if (tag === 'code' || tag === 'kbd' || tag === 'samp') next.font = 'Menlo'
  if (tag === 'sup') next.superScript = true
  if (tag === 'sub') next.subScript = true

  const children = () => (node.childNodes ?? []).flatMap((child) => inlineRuns(child, next))

  if (tag === 'a') {
    const href = node.getAttribute('href')
    const runs = children()
    if (!href || /^javascript:/i.test(href)) return runs
    return [new ExternalHyperlink({ link: href, children: runs })]
  }

  return children()
}

function cellParagraphs(cell) {
  const out = []
  for (const child of cell.childNodes ?? []) {
    if (isElement(child) && BLOCK_TAGS.has(tagOf(child))) blockInto(child, out, 0)
    else {
      const runs = inlineRuns(child)
      if (runs.length > 0) out.push(new Paragraph({ children: runs }))
    }
  }
  if (out.length === 0) out.push(new Paragraph({ children: [] }))
  return out
}

function tableBlock(node) {
  const rows = []
  const rowNodes = node.querySelectorAll('tr')
  for (const rowNode of rowNodes) {
    const cells = []
    for (const cellNode of rowNode.querySelectorAll('td,th')) {
      const header = tagOf(cellNode) === 'th'
      const children = header
        ? [new Paragraph({
          children: (cellNode.childNodes ?? []).flatMap((child) => inlineRuns(child)),
          spacing: { before: 20, after: 20 }
        })]
        : cellParagraphs(cellNode)
      cells.push(new TableCell({
        children,
        shading: header ? { fill: 'F2F2F2' } : undefined
      }))
    }
    if (cells.length > 0) rows.push(new TableRow({ children: cells }))
  }
  if (rows.length === 0) return null
  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE }
  })
}

function listInto(node, out, level, kind) {
  for (const item of node.childNodes ?? []) {
    if (!isElement(item) || tagOf(item) !== 'li') continue
    const inlineChildren = []
    const nested = []
    for (const child of item.childNodes ?? []) {
      const tag = tagOf(child)
      if (tag === 'ul' || tag === 'ol') nested.push(child)
      else inlineChildren.push(...inlineRuns(child))
    }

    const options = {
      children: inlineChildren.length > 0 ? inlineChildren : [new TextRun({ text: '' })],
      spacing: { before: 40, after: 40 }
    }
    if (kind === 'ol') {
      options.numbering = { reference: 'doc-convert-ol', level: Math.min(level, 2) }
    } else {
      options.bullet = { level: Math.min(level, 8) }
    }
    out.push(new Paragraph(options))

    for (const child of nested) {
      listInto(child, out, level + 1, tagOf(child) === 'ol' ? 'ol' : 'ul')
    }
  }
}

function blockInto(node, out, level = 0) {
  if (!isElement(node)) return

  const tag = tagOf(node)
  if (tag === 'script' || tag === 'style' || tag === 'head' || tag === 'title') return

  if (HEADINGS[tag]) {
    const runs = (node.childNodes ?? []).flatMap((child) => inlineRuns(child))
    out.push(new Paragraph({ heading: HEADINGS[tag], children: runs.length > 0 ? runs : undefined }))
    return
  }

  if (tag === 'p') {
    const runs = (node.childNodes ?? []).flatMap((child) => inlineRuns(child))
    out.push(new Paragraph({ children: runs, spacing: { after: 120 } }))
    return
  }

  if (tag === 'ul' || tag === 'ol') {
    listInto(node, out, level, tag)
    return
  }

  if (tag === 'pre') {
    const text = node.textContent ?? ''
    for (const line of text.split(/\r?\n/)) {
      out.push(new Paragraph({
        children: [new TextRun({ text: line || ' ', font: 'Menlo', size: 18 })],
        spacing: { before: 0, after: 0 }
      }))
    }
    out.push(new Paragraph({ children: [] }))
    return
  }

  if (tag === 'blockquote') {
    const runs = (node.childNodes ?? []).flatMap((child) => inlineRuns(child))
    if (runs.length > 0) {
      out.push(new Paragraph({
        children: runs,
        indent: { left: 720 },
        spacing: { before: 80, after: 80 }
      }))
    }
    return
  }

  if (tag === 'table') {
    const table = tableBlock(node)
    if (table) out.push(table)
    return
  }

  if (tag === 'hr') {
    out.push(new Paragraph({
      children: [],
      border: { bottom: { style: 'single', size: 6, color: 'CCCCCC', space: 1 } }
    }))
    return
  }

  if (tag === 'br') {
    out.push(new Paragraph({ children: [new TextRun({ break: 1 })] }))
    return
  }

  if (BLOCK_TAGS.has(tag) || tag === 'body' || tag === 'html' || tag === 'root') {
    const directInline = []
    for (const child of node.childNodes ?? []) {
      if (isWhitespaceText(child)) continue
      const childTag = tagOf(child)
      if (isElement(child) && BLOCK_TAGS.has(childTag)) {
        if (directInline.length > 0) {
          out.push(new Paragraph({ children: directInline.splice(0) }))
        }
        blockInto(child, out, level)
      } else {
        directInline.push(...inlineRuns(child))
      }
    }
    if (directInline.length > 0) out.push(new Paragraph({ children: directInline }))
    return
  }

  // Unknown inline-ish wrapper: keep its runs as one paragraph.
  const runs = (node.childNodes ?? []).flatMap((child) => inlineRuns(child))
  if (runs.length > 0) out.push(new Paragraph({ children: runs }))
}

function textBlocks(text) {
  const paragraphs = String(text ?? '').replace(/\r\n?/g, '\n').split(/\n{2,}/)
  return paragraphs.map((chunk) => {
    const lines = chunk.split('\n')
    const children = []
    lines.forEach((line, index) => {
      if (index > 0) children.push(new TextRun({ break: 1 }))
      children.push(new TextRun({ text: line }))
    })
    return new Paragraph({ children })
  })
}

function buildDocument(children) {
  if (children.length === 0) children.push(new Paragraph({ children: [] }))
  return new Document({
    numbering: {
      config: [{
        reference: 'doc-convert-ol',
        levels: [0, 1, 2].map((level) => ({
          level,
          format: LevelFormat.DECIMAL,
          text: `%${level + 1}.`,
          alignment: AlignmentType.START
        }))
      }]
    },
    styles: {
      default: {
        document: {
          run: { font: 'Calibri', size: 22 },
          paragraph: { spacing: { line: 276 } }
        }
      }
    },
    sections: [{
      properties: {
        page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } }
      },
      children
    }]
  })
}

/** Convert an HTML string into a DOCX buffer. */
export async function htmlToDocxBuffer(html) {
  const root = parse(String(html ?? ''), { blockTextElements: { script: true, noscript: true, style: true } })
  const body = root.querySelector('body') ?? root
  const children = []
  for (const child of body.childNodes ?? []) {
    if (isWhitespaceText(child)) continue
    if (isElement(child) && BLOCK_TAGS.has(tagOf(child))) blockInto(child, children)
    else {
      const runs = inlineRuns(child)
      if (runs.length > 0) children.push(new Paragraph({ children: runs }))
    }
  }
  if (children.length === 0) {
    const runs = (body.childNodes ?? []).flatMap((child) => inlineRuns(child))
    if (runs.length > 0) children.push(new Paragraph({ children: runs }))
  }
  return Packer.toBuffer(buildDocument(children))
}

/** Convert plain text into a DOCX buffer (blank line = paragraph break). */
export async function textToDocxBuffer(text) {
  return Packer.toBuffer(buildDocument(textBlocks(text)))
}

/** Convert an HTML string into readable plain text. */
export function htmlToText(html) {
  const root = parse(String(html ?? ''), { blockTextElements: { script: true, noscript: true, style: true } })
  const body = root.querySelector('body') ?? root

  const walk = (node) => {
    if (node.nodeType === 3) return node.textContent ?? node.rawText ?? ''
    if (!isElement(node)) return ''
    const tag = tagOf(node)
    if (tag === 'script' || tag === 'style' || tag === 'head' || tag === 'title') return ''
    if (tag === 'br') return '\n'
    if (tag === 'img') return node.getAttribute('alt') ?? ''
    if (tag === 'td' || tag === 'th') {
      const inner = (node.childNodes ?? []).map(walk).join('')
      return `${inner.trim()}\t`
    }
    if (tag === 'li') {
      return `- ${(node.childNodes ?? []).map(walk).join('').trim()}\n`
    }
    const inner = (node.childNodes ?? []).map(walk).join('')
    if (BLOCK_TAGS.has(tag)) return `\n${inner}\n`
    return inner
  }

  const raw = (body.childNodes ?? []).map(walk).join('')
  return raw
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.replace(/[ \t]+/g, ' ').trim())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}
