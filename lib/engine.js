/**
 * LibreOffice Kit bridge.
 *
 * The kit ships a prebuilt engine plus a Node CLI. We always run the CLI in a
 * child Node process: the Host runs under Electron, and the engine's native
 * prebuilds are matched to the plain Node ABI there. Under Electron,
 * `ELECTRON_RUN_AS_NODE=1` makes `process.execPath` run as Node (verified
 * Node 24.21.0 on this machine).
 */

import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const here = path.dirname(fileURLToPath(import.meta.url))

/** Plugin package root (one level above `lib/`). */
export const pluginRoot = path.resolve(here, '..')

let cachedStatus

function cliCandidates() {
  const override = process.env.DSH_DOC_CONVERT_CLI
  return [
    override,
    path.join(pluginRoot, 'node_modules', '@deepseek-ai', 'libreoffice-kit', 'lib', 'cli.js'),
    path.join(pluginRoot, '.tools', 'libreoffice-kit', 'lib', 'cli.js')
  ].filter(Boolean)
}

/** Locate the kit CLI once and remember whether the engine is usable. */
export function engineStatus() {
  if (cachedStatus) return cachedStatus
  for (const candidate of cliCandidates()) {
    try {
      if (existsSync(candidate)) {
        cachedStatus = { available: true, cliPath: candidate }
        return cachedStatus
      }
    } catch {
      // keep looking
    }
  }
  cachedStatus = { available: false, cliPath: null }
  return cachedStatus
}

function nodeInvocation() {
  if (process.versions.electron) {
    return { command: process.execPath, env: { ELECTRON_RUN_AS_NODE: '1' } }
  }
  return { command: process.execPath, env: {} }
}

/**
 * Run one kit CLI command and resolve its parsed JSON result.
 *
 * @param {string[]} args - CLI arguments after the script path.
 * @param {{ timeoutMs?: number, signal?: AbortSignal }} options
 */
export function runCli(args, options = {}) {
  const status = engineStatus()
  if (!status.available) {
    const error = new Error(
      'The LibreOffice engine is not installed. Reinstall the doc-convert bundle with its dependencies.'
    )
    error.code = 'ENGINE_MISSING'
    return Promise.reject(error)
  }

  const timeoutMs = Number.isFinite(options.timeoutMs) ? options.timeoutMs : 120_000
  const { command, env } = nodeInvocation()

  return new Promise((resolve, reject) => {
    let settled = false
    let timer
    let stdout = ''
    let stderr = ''

    const child = spawn(command, [status.cliPath, ...args], {
      cwd: pluginRoot,
      env: { ...process.env, ...env },
      stdio: ['ignore', 'pipe', 'pipe']
    })

    const finish = (fn, value) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      options.signal?.removeEventListener('abort', onAbort)
      fn(value)
    }

    const kill = (reason) => {
      if (settled) return
      stderr += `\n[doc-convert] ${reason}`
      try {
        child.kill('SIGKILL')
      } catch {
        // already gone
      }
    }

    const onAbort = () => kill('cancelled by caller')
    if (options.signal) {
      if (options.signal.aborted) {
        kill('cancelled before start')
      } else {
        options.signal.addEventListener('abort', onAbort, { once: true })
      }
    }

    timer = setTimeout(() => kill(`timed out after ${timeoutMs}ms`), timeoutMs)
    timer.unref?.()

    child.stdout.on('data', (chunk) => {
      stdout += chunk.toString('utf8')
    })
    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString('utf8')
    })

    child.on('error', (error) => finish(reject, error))
    child.on('close', (code) => {
      if (settled) return
      const text = stdout.trim()
      const line = text.split(/\r?\n/).filter(Boolean).pop() ?? ''
      let parsed
      try {
        parsed = line ? JSON.parse(line) : undefined
      } catch {
        parsed = undefined
      }

      if (options.signal?.aborted) {
        const error = new Error('Conversion cancelled.')
        error.code = 'ABORTED'
        return finish(reject, error)
      }

      if (parsed && parsed.code === 'failed') {
        const error = new Error(parsed.error || 'LibreOffice conversion failed.')
        error.code = 'ENGINE_FAILED'
        error.details = parsed
        return finish(reject, error)
      }

      if (code !== 0) {
        const detail = parsed?.error || stderr.trim() || text || `exit code ${code}`
        const error = new Error(`LibreOffice conversion failed: ${detail}`)
        error.code = 'ENGINE_FAILED'
        error.details = { code, stdout: text, stderr: stderr.trim() }
        return finish(reject, error)
      }

      if (!parsed) {
        const error = new Error('The LibreOffice engine returned an unreadable result.')
        error.code = 'ENGINE_PROTOCOL'
        error.details = { stdout: text, stderr: stderr.trim() }
        return finish(reject, error)
      }

      finish(resolve, parsed)
    })
  })
}

/** One office-to-office / office-to-pdf conversion through the kit. */
export async function libreOfficeConvert(inputPath, outputPath, options = {}) {
  const args = ['convert', '--input', inputPath, '--output', outputPath]
  if (options.sheet) args.push('--sheet', String(options.sheet))
  const result = await runCli(args, options)
  if (!existsSync(outputPath)) {
    const error = new Error('The engine reported success but produced no output file.')
    error.code = 'ENGINE_NO_OUTPUT'
    error.details = result
    throw error
  }
  return result
}
