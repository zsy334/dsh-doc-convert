/**
 * Format registry for the document converter.
 *
 * `TARGETS` is the single source of truth for what the tool, the HTTP API and
 * the panel advertise. Keep it in sync with `convert.js`; an advertised pair
 * with no route is a bug the smoke check catches.
 */

export const FORMATS = {
  docx: { label: 'Word (docx)', family: 'word', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  doc: { label: 'Word 97-2003 (doc)', family: 'word', mime: 'application/msword' },
  odt: { label: 'OpenDocument Text', family: 'word', mime: 'application/vnd.oasis.opendocument.text' },
  txt: { label: 'Plain text', family: 'text', mime: 'text/plain' },
  md: { label: 'Markdown', family: 'text', mime: 'text/markdown' },
  html: { label: 'HTML', family: 'text', mime: 'text/html' },
  pdf: { label: 'PDF', family: 'pdf', mime: 'application/pdf' },
  xlsx: { label: 'Excel (xlsx)', family: 'calc', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
  xls: { label: 'Excel 97-2003 (xls)', family: 'calc', mime: 'application/vnd.ms-excel' },
  ods: { label: 'OpenDocument Spreadsheet', family: 'calc', mime: 'application/vnd.oasis.opendocument.spreadsheet' },
  csv: { label: 'CSV', family: 'calc', mime: 'text/csv' },
  tsv: { label: 'TSV', family: 'calc', mime: 'text/tab-separated-values' },
  pptx: { label: 'PowerPoint (pptx)', family: 'impress', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
  ppt: { label: 'PowerPoint 97-2003 (ppt)', family: 'impress', mime: 'application/vnd.ms-powerpoint' },
  odp: { label: 'OpenDocument Presentation', family: 'impress', mime: 'application/vnd.oasis.opendocument.presentation' },
  json: { label: 'JSON', family: 'data', mime: 'application/json' }
}

/** Source extension -> the targets this build can produce. */
export const TARGETS = {
  docx: ['pdf', 'html', 'md', 'txt', 'odt'],
  doc: ['docx', 'pdf', 'html', 'md', 'txt', 'odt'],
  odt: ['docx', 'pdf', 'html', 'md', 'txt'],
  txt: ['docx', 'pdf', 'html', 'md'],
  md: ['html', 'docx', 'pdf', 'txt'],
  html: ['md', 'docx', 'pdf', 'txt'],
  pdf: ['txt', 'md'],
  xlsx: ['csv', 'pdf', 'ods', 'html', 'json'],
  xls: ['xlsx', 'csv', 'pdf', 'html', 'json'],
  ods: ['xlsx', 'csv', 'pdf', 'html', 'json'],
  csv: ['xlsx', 'pdf', 'html', 'json'],
  tsv: ['csv', 'xlsx', 'json'],
  pptx: ['pdf', 'odp'],
  ppt: ['pptx', 'pdf'],
  odp: ['pptx', 'pdf']
}

const ALIASES = {
  markdown: 'md',
  mdown: 'md',
  htm: 'html',
  text: 'txt',
  word: 'docx',
  excel: 'xlsx',
  powerpoint: 'pptx',
  presentation: 'pptx',
  spreadsheet: 'xlsx'
}

/** Normalize user/model input ("Markdown", ".MD") to a registry key. */
export function normalizeFormat(value) {
  const raw = String(value ?? '').trim().toLowerCase().replace(/^\./, '')
  return ALIASES[raw] ?? raw
}

/** Registry key for a file path, or undefined for unsupported extensions. */
export function formatOf(filePath) {
  const match = /\.([a-z0-9]+)$/i.exec(String(filePath ?? ''))
  if (!match) return undefined
  const ext = match[1].toLowerCase()
  return FORMATS[ext] ? ext : undefined
}

export function targetsFor(source) {
  return TARGETS[source] ?? []
}

export function canConvert(source, target) {
  return targetsFor(source).includes(target)
}

/** All targets any source can reach, for the tool parameter enum. */
export const ALL_TARGETS = [...new Set(Object.values(TARGETS).flat())].sort()

/** A stable, client-facing matrix (arrays of registry keys). */
export function formatMatrix() {
  return Object.fromEntries(
    Object.entries(TARGETS).map(([source, targets]) => [source, [...targets]])
  )
}
