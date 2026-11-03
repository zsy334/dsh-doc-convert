/**
 * Small JSON/HTTP surface for the document panel.
 *
 * The panel is a static Client module and cannot call Host services directly,
 * so it talks to this route table. Every path is resolved against a
 * registered workspace and rejected when it escapes that root.
 */

import { createReadStream, createWriteStream, existsSync } from 'node:fs'
import { randomUUID } from 'node:crypto'
import { once } from 'node:events'
import fs from 'node:fs/promises'
import path from 'node:path'
import { convertDocument } from './convert.js'
import { engineStatus } from './engine.js'
import { FORMATS, canConvert, formatMatrix, formatOf, normalizeFormat, targetsFor } from './formats.js'
import { assertInside, defaultOutputPath, isInside } from './paths.js'

const MAX_BODY = 256 * 1024
const MAX_UPLOAD = 200 * 1024 * 1024
const MAX_LIST = 500

async function readJson(req) {
  const chunks = []
  let size = 0
  for await (const chunk of req) {
    size += chunk.length
    if (size > MAX_BODY) throw new Error('Request body is too large.')
    chunks.push(chunk)
  }
  if (chunks.length === 0) return {}
  const text = Buffer.concat(chunks).toString('utf8')
  if (!text.trim()) return {}
  try {
    return JSON.parse(text)
  } catch {
    throw new Error('Request body must be valid JSON.')
  }
}

function sendJson(res, status, value) {
  const body = JSON.stringify(value)
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'cache-control': 'no-store',
    'content-length': Buffer.byteLength(body)
  })
  res.end(body)
}

function fail(res, status, message, code = 'DOC_CONVERT_FAILED') {
  sendJson(res, status, { ok: false, error: { code, message } })
}

function workspaceById(registry, id) {
  if (!id) throw new Error('workspaceId is required.')
  const workspace = registry.get?.(id) ?? registry.list().find((entry) => entry.id === id)
  if (!workspace) throw new Error('Unknown workspace.')
  return workspace
}

function resolveInside(root, relativePath) {
  const rel = String(relativePath ?? '').replace(/\\/g, '/')
  if (rel.split('/').some((part) => part === '..')) {
    throw new Error('Path must stay inside the workspace.')
  }
  const resolved = path.resolve(root, rel)
  return resolved
}

export function registerHttp(ctx) {
  ctx.inject(['webServer', 'workspaceRegistry'], (scope) => {
    const webServer = scope.webServer
    const registry = scope.workspaceRegistry

    const sameOrigin = (req) => {
      const origin = req.headers.origin
      if (!origin) return true
      try {
        return origin === webServer.getOrigin()
      } catch {
        return true
      }
    }

    const route = (path, handler) => scope.effect(
      () => webServer.register({ kind: 'exact', path, handler }),
      `doc-convert: ${path}`
    )

    route('/doc-convert/api/status', async (req, res) => {
      if (req.method !== 'GET') return fail(res, 405, 'Method not allowed.')
      const status = engineStatus()
      sendJson(res, 200, {
        ok: true,
        engine: {
          available: status.available,
          backend: status.available ? 'libreoffice-kit' : 'javascript-only'
        },
        formats: formatMatrix(),
        labels: Object.fromEntries(Object.entries(FORMATS).map(([key, value]) => [key, value.label]))
      })
    })

    route('/doc-convert/api/workspaces', async (req, res) => {
      if (req.method !== 'GET') return fail(res, 405, 'Method not allowed.')
      const list = registry.list().map((workspace) => ({
        id: String(workspace.id),
        path: workspace.path,
        title: workspace.title
      }))
      sendJson(res, 200, { ok: true, workspaces: list })
    })

    route('/doc-convert/api/list', async (req, res) => {
      if (req.method !== 'POST') return fail(res, 405, 'Method not allowed.')
      if (!sameOrigin(req)) return fail(res, 403, 'Cross-origin request rejected.', 'DOC_CONVERT_ORIGIN')
      try {
        const body = await readJson(req)
        const workspace = workspaceById(registry, body.workspaceId)
        const target = resolveInside(workspace.path, body.path ?? '')
        await assertInside(workspace.path, target, 'path')
        const stat = await fs.stat(target).catch(() => undefined)
        if (!stat?.isDirectory()) throw new Error('Not a directory.')

        const entries = await fs.readdir(target, { withFileTypes: true })
        const listing = []
        for (const entry of entries) {
          if (entry.name.startsWith('.')) continue
          if (listing.length >= MAX_LIST) break
          const full = path.join(target, entry.name)
          if (entry.isDirectory()) {
            listing.push({ name: entry.name, path: toRelative(workspace.path, full), dir: true })
            continue
          }
          if (!entry.isFile()) continue
          const format = formatOf(entry.name)
          if (!format) continue
          const info = await fs.stat(full).catch(() => undefined)
          listing.push({
            name: entry.name,
            path: toRelative(workspace.path, full),
            dir: false,
            size: info?.size ?? 0,
            format,
            targets: targetsFor(format)
          })
        }
        listing.sort((a, b) => (a.dir === b.dir ? a.name.localeCompare(b.name) : a.dir ? -1 : 1))

        sendJson(res, 200, {
          ok: true,
          path: toRelative(workspace.path, target),
          entries: listing
        })
      } catch (error) {
        fail(res, 400, error.message)
      }
    })

    route('/doc-convert/api/convert', async (req, res) => {
      if (req.method !== 'POST') return fail(res, 405, 'Method not allowed.')
      if (!sameOrigin(req)) return fail(res, 403, 'Cross-origin request rejected.', 'DOC_CONVERT_ORIGIN')
      try {
        const body = await readJson(req)
        const workspace = workspaceById(registry, body.workspaceId)
        const input = resolveInside(workspace.path, body.path)
        await assertInside(workspace.path, input, 'path')
        if (!existsSync(input)) throw new Error('Source file not found.')

        const source = formatOf(input)
        if (!source) throw new Error('Unsupported source file.')
        const target = normalizeFormat(body.target)
        if (!canConvert(source, target)) {
          throw new Error(`Cannot convert ${source} to ${target}.`)
        }

        const output = defaultOutputPath(input, target)
        await assertInside(workspace.path, output, 'output')

        await convertDocument({
          inputPath: input,
          target,
          outputPath: output,
          sheet: body.sheet ? String(body.sheet) : undefined,
          timeoutMs: 180_000
        })

        const stat = await fs.stat(output)
        sendJson(res, 200, {
          ok: true,
          output: {
            name: path.basename(output),
            path: toRelative(workspace.path, output),
            bytes: stat.size,
            target
          }
        })
      } catch (error) {
        fail(res, 400, error.message)
      }
    })

    route('/doc-convert/api/upload', async (req, res) => {
      if (req.method !== 'POST') return fail(res, 405, 'Method not allowed.')
      if (!sameOrigin(req)) return fail(res, 403, 'Cross-origin request rejected.', 'DOC_CONVERT_ORIGIN')
      try {
        const url = new URL(req.url ?? '/', 'http://localhost')
        const workspace = workspaceById(registry, url.searchParams.get('workspaceId'))
        const dir = resolveInside(workspace.path, url.searchParams.get('dir') ?? '')
        await assertInside(workspace.path, dir, 'path')
        const dirStat = await fs.stat(dir).catch(() => undefined)
        if (!dirStat?.isDirectory()) throw new Error('Target folder not found.')

        const name = sanitizeName(String(req.headers['x-file-name'] ?? ''))
        const finalName = await uniqueName(dir, name)
        const tmp = path.join(dir, `.dsh-upload-${randomUUID()}.part`)
        const sink = createWriteStream(tmp, { flags: 'wx' })
        let size = 0
        try {
          for await (const chunk of req) {
            size += chunk.length
            if (size > MAX_UPLOAD) throw new Error('File is too large (max 200 MB).')
            if (!sink.write(chunk)) await once(sink, 'drain')
          }
          await new Promise((resolve, reject) => sink.end((err) => (err ? reject(err) : resolve())))
        } catch (error) {
          sink.destroy()
          await fs.rm(tmp, { force: true }).catch(() => {})
          throw error
        }

        const finalPath = path.join(dir, finalName)
        await fs.rename(tmp, finalPath)
        const info = await fs.stat(finalPath)
        sendJson(res, 200, {
          ok: true,
          name: finalName,
          path: toRelative(workspace.path, finalPath),
          bytes: info.size
        })
      } catch (error) {
        fail(res, 400, error.message)
      }
    })

    route('/doc-convert/api/download', async (req, res) => {
      if (req.method !== 'GET') return fail(res, 405, 'Method not allowed.')
      try {
        const url = new URL(req.url ?? '/', 'http://localhost')
        const workspace = workspaceById(registry, url.searchParams.get('workspaceId'))
        const target = resolveInside(workspace.path, url.searchParams.get('path'))
        await assertInside(workspace.path, target, 'path')
        if (!(await isInside(workspace.path, target)) || !existsSync(target)) {
          return fail(res, 404, 'File not found.')
        }
        const stat = await fs.stat(target)
        if (!stat.isFile()) return fail(res, 404, 'File not found.')

        const format = formatOf(target)
        const mime = FORMATS[format]?.mime ?? 'application/octet-stream'
        const name = path.basename(target)
        res.writeHead(200, {
          'content-type': mime,
          'content-length': stat.size,
          'content-disposition': `attachment; filename="${asciiName(name)}"; filename*=UTF-8''${encodeURIComponent(name)}`,
          'cache-control': 'no-store'
        })
        const stream = createReadStream(target)
        stream.on('error', () => res.destroy())
        stream.pipe(res)
      } catch (error) {
        fail(res, 400, error.message)
      }
    })
  })
}

function toRelative(root, target) {
  return path.relative(root, target).split(path.sep).join('/')
}

function asciiName(name) {
  return String(name).replace(/[^\x20-\x7E]/g, '_').replace(/["\\]/g, '_')
}

function sanitizeName(raw) {
  let name = raw
  try {
    name = decodeURIComponent(raw)
  } catch {
    // keep the raw value
  }
  name = String(name).split(/[\\/]/).pop()
  name = name.replace(/[\u0000-\u001f\u007f]/g, '').trim()
  if (!name || name === '.' || name === '..') throw new Error('Invalid file name.')
  if (name.length > 200) name = name.slice(-200)
  return name
}

async function uniqueName(dir, name) {
  const ext = path.extname(name)
  const base = name.slice(0, name.length - ext.length)
  let candidate = name
  let index = 1
  while (existsSync(path.join(dir, candidate))) {
    candidate = `${base}-${index}${ext}`
    index += 1
  }
  return candidate
}
