/** Host half: registers the `convert_document` tool and the panel's HTTP API. */

import { registerHttp } from './lib/http.js'
import { registerTool } from './lib/tool.js'

export const name = 'doc-convert'
export const inject = ['tools', 'workingDirectory']

export function apply(ctx) {
  registerTool(ctx)
  registerHttp(ctx)
}
