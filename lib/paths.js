/**
 * Path containment helpers. Every write this plugin performs must land inside
 * the Session workspace (tool path) or a registered workspace (HTTP path), so
 * a conversion can never be used to write outside the user's project.
 */

import { existsSync } from 'node:fs'
import fs from 'node:fs/promises'
import path from 'node:path'

async function nearestRealPath(target) {
  let probe = path.resolve(target)
  for (;;) {
    if (existsSync(probe)) {
      try {
        return await fs.realpath(probe)
      } catch {
        return probe
      }
    }
    const parent = path.dirname(probe)
    if (parent === probe) return probe
    probe = parent
  }
}

/** True when `target` (or its nearest existing ancestor) is inside `root`. */
export async function isInside(root, target) {
  const rootReal = await nearestRealPath(root)
  const targetReal = await nearestRealPath(target)
  const relative = path.relative(rootReal, targetReal)
  return relative === '' || (!relative.startsWith('..') && !path.isAbsolute(relative))
}

export async function assertInside(root, target, label = 'path') {
  if (!(await isInside(root, target))) {
    throw new Error(`Refusing to use ${label} outside the workspace: ${target}`)
  }
}

/** A sibling output name that does not overwrite an existing file. */
export function defaultOutputPath(inputPath, extension) {
  const directory = path.dirname(inputPath)
  const base = path.basename(inputPath, path.extname(inputPath))
  const sameExtension = path.extname(inputPath).toLowerCase() === `.${extension.toLowerCase()}`
  let candidate = path.join(directory, `${base}${sameExtension ? '-converted' : ''}.${extension}`)
  let index = 1
  while (existsSync(candidate)) {
    candidate = path.join(directory, `${base}${sameExtension ? '-converted' : ''}-${index}.${extension}`)
    index += 1
  }
  return candidate
}

/** The Session's effective working directory, or the deployment fallback. */
export function workspaceRoot(ctx, exec) {
  try {
    const session = exec?.agent?.session
    if (session) return ctx.workingDirectory.get(session)
  } catch {
    // fall through
  }
  try {
    return ctx.workingDirectory.defaultDirectory
  } catch {
    return process.cwd()
  }
}
