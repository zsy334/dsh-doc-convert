/**
 * `convert_document` agent tool: the conversational entry point into the
 * same conversion service the panel uses.
 */

import { existsSync } from 'node:fs'
import fs from 'node:fs/promises'
import path from 'node:path'
import { convertDocument } from './convert.js'
import {
  ALL_TARGETS,
  canConvert,
  formatOf,
  normalizeFormat,
  targetsFor
} from './formats.js'
import { assertInside, defaultOutputPath, workspaceRoot } from './paths.js'

const MATRIX_TEXT = [
  'docx/doc/odt: pdf, html, md, txt, odt;',
  'md/html/txt: docx, pdf, html, md, txt;',
  'pdf: txt, md;',
  'xlsx/xls/ods: csv, pdf, ods, html, json;',
  'csv/tsv: xlsx, pdf, html, json;',
  'pptx/ppt/odp: pdf, pptx, odp.'
].join(' ')

export function registerTool(ctx) {
  ctx.tools.register({
    name: 'convert_document',
    description:
      'Convert a document between common formats (Word, PDF, Markdown, HTML, TXT, ' +
      'Excel/CSV, PowerPoint) and write the result into the workspace. ' +
      `Supported conversions — ${MATRIX_TEXT} ` +
      'PDF is extraction-only (text/Markdown); PDF to Word or PowerPoint layout reconstruction is not supported. ' +
      'Use this tool instead of shelling out to pandoc, soffice or libreoffice.',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        input_path: {
          type: 'string',
          description: 'Source file path, absolute or relative to the Session workspace. The extension selects the source format.'
        },
        target_format: {
          type: 'string',
          enum: ALL_TARGETS,
          description: 'Target format key, for example "pdf", "docx", "md", "html", "txt", "csv" or "json".'
        },
        output_path: {
          type: 'string',
          description: 'Optional output path inside the Session workspace. Defaults to the source directory with the target extension, avoiding overwrites.'
        },
        sheet: {
          type: 'string',
          description: 'Optional exact worksheet name when exporting a multi-sheet workbook to CSV.'
        }
      },
      required: ['input_path', 'target_format']
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          path: { type: 'string' },
          source_format: { type: 'string' },
          target_format: { type: 'string' },
          bytes: { type: 'number' },
          message: { type: 'string' }
        },
        required: ['path', 'source_format', 'target_format', 'bytes', 'message']
      },
      render(_args, value) {
        return [{
          type: 'text',
          text: value?.message ?? `Converted to ${value?.path ?? 'output'}`
        }]
      }
    },
    presentCall(args) {
      return {
        card: 'generic',
        title: `Convert ${args.input_path} → ${args.target_format}`,
        kind: 'other',
        locations: [{ path: args.input_path }]
      }
    },
    async execute(args, exec) {
      const inputPath = String(args?.input_path ?? '').trim()
      const rawTarget = String(args?.target_format ?? '').trim()
      if (!inputPath) throw new Error('input_path is required.')
      if (!rawTarget) throw new Error('target_format is required.')

      const root = workspaceRoot(ctx, exec)
      const resolvedInput = path.isAbsolute(inputPath) ? inputPath : path.resolve(root, inputPath)
      await assertInside(root, resolvedInput, 'input_path')
      if (!existsSync(resolvedInput)) throw new Error(`Input file not found: ${resolvedInput}`)

      const source = formatOf(resolvedInput)
      if (!source) throw new Error(`Unsupported source extension: ${path.extname(resolvedInput) || '(none)'}`)

      const target = normalizeFormat(rawTarget)
      if (!canConvert(source, target)) {
        throw new Error(
          `Cannot convert ${source} to ${target}. Supported targets for ${source}: ${targetsFor(source).join(', ')}.`
        )
      }

      const requestedOutput = args?.output_path ? String(args.output_path).trim() : ''
      const outputPath = requestedOutput
        ? (path.isAbsolute(requestedOutput) ? requestedOutput : path.resolve(root, requestedOutput))
        : defaultOutputPath(resolvedInput, target)
      await assertInside(root, outputPath, 'output_path')

      await convertDocument({
        inputPath: resolvedInput,
        target,
        outputPath,
        sheet: args?.sheet ? String(args.sheet) : undefined,
        signal: exec?.signal,
        timeoutMs: 180_000
      })

      const stat = await fs.stat(outputPath)
      const display = path.relative(root, outputPath) || outputPath
      return {
        path: outputPath,
        source_format: source,
        target_format: target,
        bytes: stat.size,
        message: `Converted ${path.basename(resolvedInput)} (${source}) → ${display} (${target}, ${formatBytes(stat.size)})`
      }
    }
  })
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}
