/**
 * Conversion routing.
 *
 * Two engines:
 *  - LibreOffice Kit for office formats (doc/docx/odt, xls/xlsx/ods, csv,
 *    ppt/pptx/odp) and their PDF exports.
 *  - Pure JavaScript for Markdown, HTML, TXT, CSV/TSV/JSON and PDF text
 *    extraction; DOCX output is generated directly so no Office round trip is
 *    needed for HTML/Markdown sources.
 */

import fs from 'node:fs/promises'
import { existsSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import MarkdownIt from 'markdown-it'
import { NodeHtmlMarkdown } from 'node-html-markdown'
import mammoth from 'mammoth'
import writeXlsxFile from 'write-excel-file/node'
import { htmlToDocxBuffer, htmlToText, textToDocxBuffer } from './docx.js'
import { libreOfficeConvert } from './engine.js'
import { canConvert, formatOf, normalizeFormat, targetsFor } from './formats.js'

const markdown = new MarkdownIt({ html: true, linkify: true, breaks: false })

const nhm = new NodeHtmlMarkdown({ keepStrongMarker: false, keepEmphasisMarker: false })

/** Pairs handled by one LibreOffice call. */
const KIT_PAIRS = new Set([
  'docx>pdf', 'docx>odt',
  'doc>pdf', 'doc>docx', 'doc>odt',
  'odt>docx', 'odt>pdf',
  'xlsx>pdf', 'xlsx>csv', 'xlsx>ods',
  'xls>xlsx', 'xls>csv', 'xls>pdf',
  'ods>xlsx', 'ods>csv', 'ods>pdf',
  'pptx>pdf', 'pptx>odp',
  'ppt>pptx', 'ppt>pdf',
  'odp>pptx', 'odp>pdf'
])

const HTML_SHELL = (body, title) => `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>${escapeHtml(title || 'document')}</title>
<style>
  body { font-family: -apple-system, "PingFang SC", "Microsoft YaHei", Helvetica, Arial, sans-serif; line-height: 1.65; max-width: 820px; margin: 40px auto; padding: 0 24px; color: #1f2328; }
  h1, h2, h3, h4 { line-height: 1.3; margin: 1.4em 0 0.6em; }
  h1 { font-size: 1.9em; border-bottom: 1px solid #e5e7eb; padding-bottom: 0.3em; }
  h2 { font-size: 1.5em; }
  pre { background: #f6f8fa; padding: 12px 14px; overflow: auto; border-radius: 6px; }
  code { font-family: Menlo, Consolas, monospace; background: #f6f8fa; padding: 0.15em 0.35em; border-radius: 4px; }
  pre code { padding: 0; background: transparent; }
  blockquote { margin: 1em 0; padding: 0.2em 1em; color: #57606a; border-left: 4px solid #d0d7de; }
  table { border-collapse: collapse; margin: 1em 0; }
  th, td { border: 1px solid #d0d7de; padding: 6px 10px; }
  img { max-width: 100%; }
</style>
</head>
<body>
${body}
</body>
</html>`

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

async function writeText(file, text) {
  await fs.mkdir(path.dirname(file), { recursive: true })
  await fs.writeFile(file, text, 'utf8')
}

async function writeBuffer(file, buffer) {
  await fs.mkdir(path.dirname(file), { recursive: true })
  await fs.writeFile(file, buffer)
}

function mdToHtmlBody(source) {
  return markdown.render(String(source ?? ''))
}

function htmlToMarkdown(html) {
  return nhm.translate(String(html ?? '')).replace(/\n{3,}/g, '\n\n').trim() + '\n'
}

async function docxToHtml(file) {
  const result = await mammoth.convertToHtml({ path: file })
  return result.value
}

async function tsvToCsv(input, output) {
  const text = await fs.readFile(input, 'utf8')
  const delimiter = guessDelimiter(text, '\t')
  const rows = parseDelimited(text, delimiter)
  await writeText(output, stringifyDelimited(rows, ','))
}

function guessDelimiter(text, fallback) {
  const firstLine = String(text ?? '').split(/\r?\n/, 1)[0] ?? ''
  if (fallback === '\t') return '\t'
  const commas = (firstLine.match(/,/g) ?? []).length
  const semis = (firstLine.match(/;/g) ?? []).length
  if (semis > commas) return ';'
  return fallback
}

export function parseDelimited(text, delimiter = ',') {
  const source = String(text ?? '').replace(/^\uFEFF/, '')
  const rows = []
  let row = []
  let field = ''
  let quoted = false

  for (let index = 0; index < source.length; index += 1) {
    const char = source[index]
    if (quoted) {
      if (char === '"') {
        if (source[index + 1] === '"') {
          field += '"'
          index += 1
        } else {
          quoted = false
        }
      } else {
        field += char
      }
      continue
    }
    if (char === '"') {
      quoted = true
    } else if (char === delimiter) {
      row.push(field)
      field = ''
    } else if (char === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else if (char === '\r') {
      // newline is handled by \n
    } else {
      field += char
    }
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field)
    rows.push(row)
  }
  return rows.filter((entry) => entry.some((cell) => cell.length > 0))
}

export function stringifyDelimited(rows, delimiter = ',') {
  return rows.map((row) => row.map((cell) => {
    const value = cell == null ? '' : String(cell)
    return new RegExp(`["\\n\\r${delimiter === '\t' ? '\\t' : delimiter}]`).test(value)
      ? `"${value.replace(/"/g, '""')}"`
      : value
  }).join(delimiter)).join('\n')
}

function csvToJson(input) {
  const text = input
  const rows = parseDelimited(text, guessDelimiter(text, ','))
  if (rows.length === 0) return '[]\n'
  const headers = rows[0].map((header, index) => header || `column_${index + 1}`)
  const objects = rows.slice(1).map((row) => Object.fromEntries(
    headers.map((header, index) => [header, row[index] ?? ''])
  ))
  return `${JSON.stringify(objects, null, 2)}\n`
}

function csvToHtml(input) {
  const rows = parseDelimited(input, guessDelimiter(input, ','))
  if (rows.length === 0) return HTML_SHELL('', 'table')
  const [head, ...body] = rows
  const headHtml = `<thead><tr>${head.map((cell) => `<th>${escapeHtml(cell)}</th>`).join('')}</tr></thead>`
  const bodyHtml = `<tbody>${body.map((row) => `<tr>${head.map((_, index) => `<td>${escapeHtml(row[index] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody>`
  return HTML_SHELL(`<table>${headHtml}${bodyHtml}</table>`, 'table')
}

async function pdfToText(input) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const data = new Uint8Array(await fs.readFile(input))
  const document = await pdfjs.getDocument({
    data,
    useSystemFonts: true,
    isEvalSupported: false,
    verbosity: 0
  }).promise
  const pages = []
  try {
    for (let index = 1; index <= document.numPages; index += 1) {
      const page = await document.getPage(index)
      const content = await page.getTextContent()
      let text = ''
      for (const item of content.items) {
        if (!('str' in item)) continue
        text += item.str
        text += item.hasEOL ? '\n' : ' '
      }
      pages.push(text.replace(/[ \t]+\n/g, '\n').replace(/[ \t]{2,}/g, ' ').trim())
    }
  } finally {
    await document.destroy()
  }
  return pages.join('\n\n').trim()
}

/**
 * Convert one file. The caller owns path validation and output placement.
 *
 * @param {{
 *   inputPath: string,
 *   target: string,
 *   outputPath: string,
 *   sheet?: string,
 *   signal?: AbortSignal,
 *   timeoutMs?: number
 * }} request
 */
export async function convertDocument(request) {
  const source = formatOf(request.inputPath)
  const target = normalizeFormat(request.target)
  if (!source) {
    throw new Error(`Unsupported source format: ${path.extname(request.inputPath) || '(no extension)'}`)
  }
  if (!canConvert(source, target)) {
    throw new Error(`Cannot convert ${source} to ${target}. Supported targets: ${targetsFor(source).join(', ') || 'none'}.`)
  }
  if (path.resolve(request.inputPath) === path.resolve(request.outputPath)) {
    throw new Error('The output path is the input path; choose a different file name or format.')
  }

  const work = await fs.mkdtemp(path.join(os.tmpdir(), 'dsh-doc-convert-'))
  const ctx = { ...request, work }
  const outputPath = request.outputPath
  // Produce inside the private work directory, then publish. LibreOffice
  // refuses to overwrite an existing file, and this keeps a failed conversion
  // from truncating a previous output.
  const staged = path.join(work, `result.${target}`)

  try {
    await produce(request.inputPath, source, target, staged, ctx)
    if (!existsSync(staged)) {
      throw new Error('Conversion finished without producing an output file.')
    }
    const stat = await fs.stat(staged)
    if (!stat.isFile() || stat.size === 0) {
      throw new Error('Conversion produced an empty file.')
    }
    await fs.mkdir(path.dirname(outputPath), { recursive: true })
    await fs.copyFile(staged, outputPath)
    return { outputPath, bytes: stat.size }
  } finally {
    await fs.rm(work, { recursive: true, force: true }).catch(() => {})
  }
}

async function intermediate(ctx, extension) {
  const file = path.join(ctx.work, `step-${Math.random().toString(36).slice(2, 10)}.${extension}`)
  return file
}

async function lo(input, output, ctx) {
  await libreOfficeConvert(input, output, {
    sheet: ctx.sheet,
    signal: ctx.signal,
    timeoutMs: ctx.timeoutMs ?? 120_000
  })
}

async function produce(input, from, target, output, ctx) {
  const key = `${from}>${target}`

  if (KIT_PAIRS.has(key)) {
    await lo(input, output, ctx)
    return
  }

  switch (key) {
    // ---------- Word family ----------
    case 'docx>html':
      await writeText(output, HTML_SHELL(await docxToHtml(input), path.basename(input)))
      return
    case 'docx>md':
      await writeText(output, htmlToMarkdown(await docxToHtml(input)))
      return
    case 'docx>txt':
      await writeText(output, htmlToText(await docxToHtml(input)))
      return
    case 'doc>html':
    case 'doc>md':
    case 'doc>txt': {
      const docx = await intermediate(ctx, 'docx')
      await lo(input, docx, ctx)
      await produce(docx, 'docx', target, output, ctx)
      return
    }
    case 'odt>html':
    case 'odt>md':
    case 'odt>txt': {
      const docx = await intermediate(ctx, 'docx')
      await lo(input, docx, ctx)
      await produce(docx, 'docx', target, output, ctx)
      return
    }

    // ---------- Markdown ----------
    case 'md>html': {
      const body = mdToHtmlBody(await fs.readFile(input, 'utf8'))
      await writeText(output, HTML_SHELL(body, path.basename(input)))
      return
    }
    case 'md>txt': {
      const body = mdToHtmlBody(await fs.readFile(input, 'utf8'))
      await writeText(output, htmlToText(body))
      return
    }
    case 'md>docx': {
      const body = mdToHtmlBody(await fs.readFile(input, 'utf8'))
      await writeBuffer(output, await htmlToDocxBuffer(body))
      return
    }
    case 'md>pdf': {
      const body = mdToHtmlBody(await fs.readFile(input, 'utf8'))
      const docx = await intermediate(ctx, 'docx')
      await writeBuffer(docx, await htmlToDocxBuffer(body))
      await lo(docx, output, ctx)
      return
    }

    // ---------- HTML ----------
    case 'html>md':
      await writeText(output, htmlToMarkdown(await fs.readFile(input, 'utf8')))
      return
    case 'html>txt':
      await writeText(output, htmlToText(await fs.readFile(input, 'utf8')))
      return
    case 'html>docx':
      await writeBuffer(output, await htmlToDocxBuffer(await fs.readFile(input, 'utf8')))
      return
    case 'html>pdf': {
      const docx = await intermediate(ctx, 'docx')
      await writeBuffer(docx, await htmlToDocxBuffer(await fs.readFile(input, 'utf8')))
      await lo(docx, output, ctx)
      return
    }

    // ---------- Plain text ----------
    case 'txt>md':
      await writeText(output, await fs.readFile(input, 'utf8'))
      return
    case 'txt>html': {
      const text = await fs.readFile(input, 'utf8')
      const body = text.replace(/\r\n?/g, '\n').split(/\n{2,}/).map((block) => {
        const lines = block.split('\n').map((line) => escapeHtml(line)).join('<br>')
        return `<p>${lines}</p>`
      }).join('\n')
      await writeText(output, HTML_SHELL(body, path.basename(input)))
      return
    }
    case 'txt>docx':
      await writeBuffer(output, await textToDocxBuffer(await fs.readFile(input, 'utf8')))
      return
    case 'txt>pdf': {
      const docx = await intermediate(ctx, 'docx')
      await writeBuffer(docx, await textToDocxBuffer(await fs.readFile(input, 'utf8')))
      await lo(docx, output, ctx)
      return
    }

    // ---------- PDF ----------
    case 'pdf>txt': {
      await writeText(output, await pdfToText(input))
      return
    }
    case 'pdf>md': {
      const text = await pdfToText(input)
      const md = text.split(/\n{2,}/).map((paragraph) => paragraph.replace(/\n/g, ' ')).join('\n\n')
      await writeText(output, md)
      return
    }

    // ---------- Spreadsheets ----------
    case 'xlsx>html':
    case 'xls>html':
    case 'ods>html': {
      const csv = await intermediate(ctx, 'csv')
      await produce(input, from, 'csv', csv, ctx)
      await writeText(output, csvToHtml(await fs.readFile(csv, 'utf8')))
      return
    }
    case 'xlsx>json':
    case 'xls>json':
    case 'ods>json': {
      const csv = await intermediate(ctx, 'csv')
      await produce(input, from, 'csv', csv, ctx)
      await writeText(output, csvToJson(await fs.readFile(csv, 'utf8')))
      return
    }
    case 'csv>html':
      await writeText(output, csvToHtml(await fs.readFile(input, 'utf8')))
      return
    case 'csv>json':
      await writeText(output, csvToJson(await fs.readFile(input, 'utf8')))
      return
    case 'csv>xlsx': {
      const text = await fs.readFile(input, 'utf8')
      const rows = parseDelimited(text, guessDelimiter(text, ','))
      await writeXlsxFile(rows).toFile(output)
      return
    }
    case 'csv>pdf': {
      const xlsx = await intermediate(ctx, 'xlsx')
      await produce(input, 'csv', 'xlsx', xlsx, ctx)
      await lo(xlsx, output, ctx)
      return
    }
    case 'tsv>xlsx': {
      const csv = await intermediate(ctx, 'csv')
      await tsvToCsv(input, csv)
      await produce(csv, 'csv', 'xlsx', output, ctx)
      return
    }
    case 'tsv>csv':
      await tsvToCsv(input, output)
      return
    case 'tsv>json': {
      const csvFile = await intermediate(ctx, 'csv')
      await tsvToCsv(input, csvFile)
      await writeText(output, csvToJson(await fs.readFile(csvFile, 'utf8')))
      return
    }

    default: {
      throw new Error(`Conversion route ${key} is not implemented yet.`)
    }
  }
}
