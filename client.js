window.__ModuleLoader__.load({
  id: 'dsh-doc-convert',
  factory(require) {
    const React = require('react')
    const h = React.createElement
    const { useCallback, useEffect, useMemo, useState } = React

    const ZH = {
      panelTitle: '文档转换',
      subtitle: '在工作区文件之间转换 Word、PDF、Markdown、HTML、TXT、Excel/CSV 与 PowerPoint。',
      engineReady: 'LibreOffice 引擎就绪',
      engineMissing: '仅轻量引擎',
      engineMissingHint: 'Office 与 PDF 版式转换需要重装插件依赖。',
      workspace: '工作区',
      refresh: '刷新',
      files: '文件',
      emptyDir: '这个目录里没有可转换的文件。',
      selectFile: '从左侧选择文件后设置目标格式。',
      noTargets: '该格式没有可用的转换目标。',
      targetFormat: '目标格式',
      convert: '开始转换',
      converting: '正在转换…',
      result: '转换完成',
      download: '下载文件',
      size: '大小',
      up: '上一级',
      error: '转换失败',
      supported: '支持：Word (docx/doc/odt)、PDF、Markdown、HTML、TXT、Excel (xlsx/xls/ods)、CSV/TSV/JSON、PowerPoint (pptx/ppt/odp)。',
      hint: '把文件放进工作区后，在这里选择它并转换；结果会写回同一目录。',
      loading: '加载中…',
      source: '源格式',
      destination: '目标格式',
      sheetLabel: '工作表（可选）',
      sheetPlaceholder: '多工作表时填写准确名称，例如 Sheet1',
      upload: '选择文件',
      dropHint: '把文件拖到这里，或点击“选择文件”从本机上传',
      uploading: '上传中…'
    }

    const EN = {
      panelTitle: 'Document Converter',
      subtitle: 'Convert Word, PDF, Markdown, HTML, TXT, Excel/CSV and PowerPoint files in your workspace.',
      engineReady: 'LibreOffice engine ready',
      engineMissing: 'Lite engine only',
      engineMissingHint: 'Reinstall the plugin dependencies for Office and PDF layout conversion.',
      workspace: 'Workspace',
      refresh: 'Refresh',
      files: 'Files',
      emptyDir: 'No convertible files in this folder.',
      selectFile: 'Pick a file on the left, then choose a target format.',
      noTargets: 'This format has no available target.',
      targetFormat: 'Target format',
      convert: 'Convert',
      converting: 'Converting…',
      result: 'Conversion complete',
      download: 'Download',
      size: 'Size',
      up: 'Up',
      error: 'Conversion failed',
      supported: 'Word (docx/doc/odt), PDF, Markdown, HTML, TXT, Excel (xlsx/xls/ods), CSV/TSV/JSON, PowerPoint (pptx/ppt/odp).',
      hint: 'Place files in the workspace, pick one here, and convert; output lands in the same folder.',
      loading: 'Loading…',
      source: 'Source',
      destination: 'Target',
      sheetLabel: 'Sheet (optional)',
      sheetPlaceholder: 'Exact sheet name for multi-sheet workbooks',
      upload: 'Choose files',
      dropHint: 'Drop files here, or click "Choose files" to upload from this computer',
      uploading: 'Uploading…'
    }

    function dictionary(locale) {
      return locale === 'en' ? EN : ZH
    }

    async function api(path, init) {
      const response = await fetch(path, {
        headers: { 'content-type': 'application/json' },
        ...init
      })
      const data = await response.json().catch(() => undefined)
      if (!response.ok || !data || data.ok === false) {
        throw new Error(data?.error?.message || `HTTP ${response.status}`)
      }
      return data
    }

    function formatBytes(bytes) {
      if (!Number.isFinite(bytes) || bytes <= 0) return '—'
      if (bytes < 1024) return `${bytes} B`
      if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
      return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
    }

    function Icon({ size = 18 }) {
      return h('svg', {
        viewBox: '0 0 24 24',
        width: size,
        height: size,
        fill: 'none',
        stroke: 'currentColor',
        strokeWidth: 1.7,
        strokeLinecap: 'round',
        strokeLinejoin: 'round',
        'aria-hidden': true,
        style: { display: 'block' }
      },
      h('path', { d: 'M7 3.5h6.5L18 8v12a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1v-15a1 1 0 0 1 1-1z' }),
      h('path', { d: 'M13.5 3.5V8H18' }),
      h('path', { d: 'M9.5 12.5h5' }),
      h('path', { d: 'M9.5 16h3' }))
    }

    const PANEL_CSS = `
        .dcv-page { max-width: 1100px; margin: 0 auto; padding: 26px 30px 40px; color: var(--dsw-alias-label-primary); font-size: 13px; }
        .dcv-header { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; margin-bottom: 18px; }
        .dcv-title { margin: 0 0 4px; font-size: 20px; font-weight: 600; }
        .dcv-subtitle { margin: 0; color: var(--dsw-alias-label-secondary); line-height: 1.6; max-width: 680px; }
        .dcv-header-side { display: flex; align-items: center; gap: 10px; flex-shrink: 0; }
        .dcv-badge { padding: 3px 9px; border-radius: 999px; font-size: 11px; border: 1px solid var(--dsw-alias-border-l1); color: var(--dsw-alias-label-secondary); white-space: nowrap; }
        .dcv-badge.ok { color: var(--dsw-alias-state-success-primary); border-color: var(--dsw-alias-state-success-primary); }
        .dcv-badge.warn { color: var(--dsw-alias-state-warn-primary); border-color: var(--dsw-alias-state-warn-primary); }
        .dcv-toolbar { display: flex; align-items: center; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; }
        .dcv-field { display: flex; align-items: center; gap: 8px; }
        .dcv-field-label { color: var(--dsw-alias-label-secondary); font-size: 12px; }
        .dcv-select { background: var(--dsw-alias-bg-layer-1); color: var(--dsw-alias-label-primary); border: 1px solid var(--dsw-alias-border-l1); border-radius: 7px; padding: 6px 10px; font-size: 13px; }
        .dcv-sheet-field { display: flex; flex-direction: column; gap: 5px; }
        .dcv-input { background: var(--dsw-alias-bg-base); color: var(--dsw-alias-label-primary); border: 1px solid var(--dsw-alias-border-l1); border-radius: 7px; padding: 6px 10px; font-size: 13px; }
        .dcv-path { color: var(--dsw-alias-label-secondary); font-size: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 520px; }
        .dcv-drop { border: 1px dashed var(--dsw-alias-border-l2); border-radius: 10px; padding: 14px 16px; margin-bottom: 14px; color: var(--dsw-alias-label-secondary); text-align: center; font-size: 12px; background: var(--dsw-alias-bg-layer-1); transition: border-color 120ms, color 120ms; }
        .dcv-drop.active { border-color: var(--dsw-alias-brand-primary); color: var(--dsw-alias-brand-primary); background: var(--dsw-alias-bg-layer-2); }
        .dcv-grid { display: grid; grid-template-columns: minmax(0, 1.15fr) minmax(0, 1fr); gap: 16px; align-items: start; }
        @media (max-width: 900px) { .dcv-grid { grid-template-columns: minmax(0, 1fr); } }
        .dcv-card { background: var(--dsw-alias-bg-layer-1); border: 1px solid var(--dsw-alias-border-l1); border-radius: 10px; padding: 14px; min-width: 0; }
        .dcv-side { display: flex; flex-direction: column; gap: 12px; }
        .dcv-card-head { display: flex; align-items: center; justify-content: space-between; gap: 8px; margin-bottom: 8px; }
        .dcv-card-title { font-weight: 600; font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .dcv-hint { color: var(--dsw-alias-label-secondary); font-size: 12px; line-height: 1.6; }
        .dcv-crumbs { display: flex; align-items: center; gap: 2px; flex-wrap: wrap; margin-bottom: 6px; }
        .dcv-crumb { background: none; border: none; color: var(--dsw-alias-brand-primary); cursor: pointer; padding: 2px 4px; font-size: 12px; }
        .dcv-crumb-sep { color: var(--dsw-alias-label-secondary); }
        .dcv-up { background: none; border: none; color: var(--dsw-alias-label-secondary); cursor: pointer; padding: 4px 2px; font-size: 12px; margin-bottom: 4px; }
        .dcv-list { list-style: none; margin: 0; padding: 0; max-height: 420px; overflow: auto; }
        .dcv-row { width: 100%; display: flex; align-items: center; justify-content: space-between; gap: 10px; background: none; border: 1px solid transparent; border-radius: 7px; padding: 7px 9px; color: var(--dsw-alias-label-primary); cursor: pointer; text-align: left; font-size: 13px; }
        .dcv-row:hover { background: var(--dsw-alias-bg-layer-2); }
        .dcv-row.selected { border-color: var(--dsw-alias-brand-primary); background: var(--dsw-alias-bg-layer-2); }
        .dcv-row-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .dcv-row-meta { color: var(--dsw-alias-label-secondary); font-size: 11px; white-space: nowrap; }
        .dcv-empty { color: var(--dsw-alias-label-secondary); margin: 8px 0; line-height: 1.6; }
        .dcv-placeholder { display: flex; flex-direction: column; gap: 8px; }
        .dcv-meta { display: flex; gap: 24px; margin: 0; }
        .dcv-meta dt { color: var(--dsw-alias-label-secondary); font-size: 11px; margin-bottom: 2px; }
        .dcv-meta dd { margin: 0; font-size: 13px; }
        .dcv-formats { display: flex; flex-direction: column; gap: 8px; }
        .dcv-chips { display: flex; flex-wrap: wrap; gap: 7px; }
        .dcv-chip { border: 1px solid var(--dsw-alias-border-l1); background: var(--dsw-alias-bg-base); color: var(--dsw-alias-label-primary); border-radius: 999px; padding: 4px 12px; font-size: 12px; cursor: pointer; }
        .dcv-chip:hover { border-color: var(--dsw-alias-border-l2); }
        .dcv-chip.active { border-color: var(--dsw-alias-brand-primary); color: var(--dsw-alias-brand-primary); }
        .dcv-btn { border: 1px solid var(--dsw-alias-border-l1); background: var(--dsw-alias-bg-layer-2); color: var(--dsw-alias-label-primary); border-radius: 8px; padding: 7px 14px; font-size: 12px; cursor: pointer; text-decoration: none; display: inline-flex; align-items: center; justify-content: center; gap: 6px; }
        .dcv-btn:hover { border-color: var(--dsw-alias-border-l2); }
        .dcv-btn:disabled { opacity: 0.55; cursor: default; }
        .dcv-btn.primary { background: var(--dsw-alias-brand-primary); border-color: var(--dsw-alias-brand-primary); color: #fff; }
        .dcv-result { border-top: 1px solid var(--dsw-alias-border-l1); padding-top: 12px; display: flex; flex-direction: column; gap: 8px; align-items: flex-start; }
        .dcv-result-title { color: var(--dsw-alias-state-success-primary); font-weight: 600; }
        .dcv-result-name { font-size: 13px; word-break: break-all; }
        .dcv-error { border: 1px solid var(--dsw-alias-state-error-primary); color: var(--dsw-alias-state-error-primary); border-radius: 8px; padding: 8px 12px; margin-bottom: 12px; font-size: 12px; line-height: 1.6; }
        .dcv-footer-hint { margin-top: 14px; color: var(--dsw-alias-label-secondary); font-size: 11px; }
      `

    function makePanel(ctx) {
      return function Panel() {
        const [, setLocaleTick] = useState(0)
        const locale = ctx.locale.getSnapshot().active
        const t = dictionary(locale)

        const [status, setStatus] = useState(null)
        const [workspaces, setWorkspaces] = useState([])
        const [workspaceId, setWorkspaceId] = useState('')
        const [directory, setDirectory] = useState('')
        const [entries, setEntries] = useState([])
        const [selected, setSelected] = useState(null)
        const [target, setTarget] = useState('')
        const [sheet, setSheet] = useState('')
        const [busy, setBusy] = useState(false)
        const [loading, setLoading] = useState(true)
        const [error, setError] = useState('')
        const [result, setResult] = useState(null)
        const [dragging, setDragging] = useState(false)
        const [uploading, setUploading] = useState(false)

        useEffect(() => {
          const off = ctx.locale.subscribe(() => setLocaleTick((value) => value + 1))
          return () => off?.()
        }, [])

        useEffect(() => {
          let cancelled = false
          ;(async () => {
            try {
              const [statusData, workspaceData] = await Promise.all([
                api('/doc-convert/api/status'),
                api('/doc-convert/api/workspaces')
              ])
              if (cancelled) return
              setStatus(statusData)
              setWorkspaces(workspaceData.workspaces)
              setWorkspaceId((current) => current || workspaceData.workspaces[0]?.id || '')
            } catch (loadError) {
              if (!cancelled) setError(loadError.message)
            } finally {
              if (!cancelled) setLoading(false)
            }
          })()
          return () => { cancelled = true }
        }, [])

        const listDirectory = useCallback(async (id, dirPath) => {
          if (!id) return
          setLoading(true)
          setError('')
          try {
            const data = await api('/doc-convert/api/list', {
              method: 'POST',
              body: JSON.stringify({ workspaceId: id, path: dirPath })
            })
            setDirectory(data.path)
            setEntries(data.entries)
            setSelected(null)
            setTarget('')
            setResult(null)
          } catch (listError) {
            setError(listError.message)
            setEntries([])
          } finally {
            setLoading(false)
          }
        }, [])

        useEffect(() => {
          if (workspaceId) listDirectory(workspaceId, '')
        }, [workspaceId, listDirectory])

        const workspace = useMemo(
          () => workspaces.find((entry) => entry.id === workspaceId),
          [workspaces, workspaceId]
        )

        const targets = selected?.targets ?? []
        const crumbs = directory ? directory.split('/').filter(Boolean) : []

        const pick = (entry) => {
          setSelected(entry)
          setTarget(entry.targets?.[0] ?? '')
          setSheet('')
          setResult(null)
          setError('')
        }

        const convert = async () => {
          if (!selected || !target || busy) return
          setBusy(true)
          setError('')
          setResult(null)
          try {
            const data = await api('/doc-convert/api/convert', {
              method: 'POST',
              body: JSON.stringify({ workspaceId, path: selected.path, target, sheet: sheet.trim() || undefined })
            })
            setResult(data.output)
          } catch (convertError) {
            setError(convertError.message)
          } finally {
            setBusy(false)
          }
        }

        const refresh = () => listDirectory(workspaceId, directory)

        const uploadOne = async (file) => {
          const url = `/doc-convert/api/upload?workspaceId=${encodeURIComponent(workspaceId)}&dir=${encodeURIComponent(directory)}`
          const response = await fetch(url, {
            method: 'POST',
            headers: { 'x-file-name': encodeURIComponent(file.name) },
            body: file
          })
          const data = await response.json().catch(() => undefined)
          if (!response.ok || !data || data.ok === false) {
            throw new Error(data?.error?.message || `Upload failed (HTTP ${response.status})`)
          }
          return data
        }

        const uploadFiles = async (files) => {
          if (!files.length || !workspaceId) return
          setUploading(true)
          setError('')
          try {
            for (const file of files) {
              await uploadOne(file)
            }
          } catch (uploadError) {
            setError(uploadError.message)
          } finally {
            setUploading(false)
            listDirectory(workspaceId, directory)
          }
        }

        const onPickFiles = (event) => {
          const files = Array.from(event.target.files ?? [])
          event.target.value = ''
          uploadFiles(files)
        }

        const downloadUrl = result
          ? `/doc-convert/api/download?workspaceId=${encodeURIComponent(workspaceId)}&path=${encodeURIComponent(result.path)}`
          : '#'

        const engineReady = status?.engine?.available === true

        return h(React.Fragment, null,
          h('style', null, PANEL_CSS),
          h('div', {
            className: 'dcv-page' + (dragging ? ' dragging' : ''),
            onDragOver: (event) => { event.preventDefault(); setDragging(true) },
            onDragLeave: (event) => { if (!event.currentTarget.contains(event.relatedTarget)) setDragging(false) },
            onDrop: (event) => {
              event.preventDefault()
              setDragging(false)
              uploadFiles(Array.from(event.dataTransfer?.files ?? []))
            }
          },
            h('header', { className: 'dcv-header' },
              h('div', null,
                h('h1', { className: 'dcv-title' }, t.panelTitle),
                h('p', { className: 'dcv-subtitle' }, t.subtitle)
              ),
              h('div', { className: 'dcv-header-side' },
                h('span', {
                  className: 'dcv-badge ' + (engineReady ? 'ok' : 'warn'),
                  title: engineReady ? 'libreoffice-kit' : t.engineMissingHint
                }, engineReady ? t.engineReady : t.engineMissing),
                h('button', { className: 'dcv-btn', onClick: refresh, disabled: loading }, t.refresh),
                h('label', { className: 'dcv-btn' },
                  h('input', { type: 'file', multiple: true, style: { display: 'none' }, onChange: onPickFiles }),
                  t.upload)
              )
            ),

            h('div', { className: 'dcv-toolbar' },
              h('label', { className: 'dcv-field' },
                h('span', { className: 'dcv-field-label' }, t.workspace),
                h('select', {
                  className: 'dcv-select',
                  value: workspaceId,
                  onChange: (event) => setWorkspaceId(event.target.value)
                }, workspaces.map((entry) => h('option', { key: entry.id, value: entry.id }, entry.title || entry.path)))
              ),
              workspace ? h('code', { className: 'dcv-path' }, workspace.path) : null
            ),

            h('div', { className: 'dcv-drop' + (dragging ? ' active' : '') },
              uploading ? t.uploading : t.dropHint),

            error ? h('div', { className: 'dcv-error', role: 'alert' }, h('strong', null, t.error), ' · ', error) : null,

            h('div', { className: 'dcv-grid' },
              h('section', { className: 'dcv-card' },
                h('div', { className: 'dcv-card-head' },
                  h('span', { className: 'dcv-card-title' }, t.files),
                  h('span', { className: 'dcv-hint' }, loading ? t.loading : `${entries.length}`)
                ),
                h('div', { className: 'dcv-crumbs' },
                  h('button', { className: 'dcv-crumb', onClick: () => listDirectory(workspaceId, '') }, 'root'),
                  crumbs.map((crumb, index) => h('span', { key: index, className: 'dcv-crumb-wrap' },
                    h('span', { className: 'dcv-crumb-sep' }, '/'),
                    h('button', {
                      className: 'dcv-crumb',
                      onClick: () => listDirectory(workspaceId, crumbs.slice(0, index + 1).join('/'))
                    }, crumb)
                  ))
                ),
                directory ? h('button', {
                  className: 'dcv-up',
                  onClick: () => listDirectory(workspaceId, crumbs.slice(0, -1).join('/'))
                }, '↑ ' + t.up) : null,
                entries.length === 0 && !loading
                  ? h('p', { className: 'dcv-empty' }, t.emptyDir)
                  : h('ul', { className: 'dcv-list' }, entries.map((entry) => h('li', { key: entry.path },
                    h('button', {
                      className: 'dcv-row' + (selected?.path === entry.path ? ' selected' : '') + (entry.dir ? ' dir' : ''),
                      onClick: () => (entry.dir ? listDirectory(workspaceId, entry.path) : pick(entry))
                    },
                    h('span', { className: 'dcv-row-name' }, entry.dir ? '📁 ' : '📄 ', entry.name),
                    h('span', { className: 'dcv-row-meta' },
                      entry.dir ? '' : `${entry.format} · ${formatBytes(entry.size)}`)
                    )
                  )))
              ),

              h('section', { className: 'dcv-card dcv-side' },
                !selected
                  ? h('div', { className: 'dcv-placeholder' },
                    h('p', { className: 'dcv-empty' }, t.selectFile),
                    h('p', { className: 'dcv-hint' }, t.hint),
                    h('p', { className: 'dcv-hint' }, t.supported))
                  : h(React.Fragment, null,
                    h('div', { className: 'dcv-card-head' },
                      h('span', { className: 'dcv-card-title' }, selected.name)
                    ),
                    h('dl', { className: 'dcv-meta' },
                      h('div', null, h('dt', null, t.source), h('dd', null, (status?.labels?.[selected.format] ?? selected.format))),
                      h('div', null, h('dt', null, t.size), h('dd', null, formatBytes(selected.size)))
                    ),
                    h('div', { className: 'dcv-formats' },
                      h('span', { className: 'dcv-field-label' }, t.targetFormat),
                      targets.length === 0
                        ? h('p', { className: 'dcv-hint' }, t.noTargets)
                        : h('div', { className: 'dcv-chips' }, targets.map((key) => h('button', {
                          key,
                          className: 'dcv-chip' + (target === key ? ' active' : ''),
                          onClick: () => { setTarget(key); setResult(null) }
                        }, status?.labels?.[key] ?? key)))
                    ),
                    ['xlsx', 'xls', 'ods'].includes(selected.format) && target === 'csv'
                      ? h('label', { className: 'dcv-sheet-field' },
                        h('span', { className: 'dcv-field-label' }, t.sheetLabel),
                        h('input', {
                          className: 'dcv-input',
                          value: sheet,
                          placeholder: t.sheetPlaceholder,
                          onChange: (event) => setSheet(event.target.value)
                        }))
                      : null,
                    h('button', {
                      className: 'dcv-btn primary',
                      disabled: !target || busy,
                      onClick: convert
                    }, busy ? t.converting : t.convert),

                    result
                      ? h('div', { className: 'dcv-result' },
                        h('div', { className: 'dcv-result-title' }, '✓ ' + t.result),
                        h('div', { className: 'dcv-result-name' }, result.name),
                        h('div', { className: 'dcv-hint' }, `${result.target} · ${formatBytes(result.bytes)}`),
                        h('a', { className: 'dcv-btn primary', href: downloadUrl, download: result.name }, t.download),
                        h('button', {
                          className: 'dcv-btn',
                          onClick: () => listDirectory(workspaceId, directory)
                        }, t.refresh)
                      )
                      : null
                  )
              )
            ),

            h('p', { className: 'dcv-footer-hint' }, t.hint)
          )
        )
      }
    }

    return {
      inject: ['slots', 'locale'],
      apply(ctx) {
        ctx.locale.register('doc-convert', { en: EN, zh: ZH })
        const label = () => dictionary(ctx.locale.getSnapshot().active).panelTitle

        ctx.slots.inject('sidebar.panellist', () => ctx.slots.register({
          name: 'sidebar.panellist',
          id: 'doc-convert',
          order: 20,
          label
        }, Icon))

        ctx.slots.inject('main', () => ctx.slots.register({
          name: 'main',
          key: 'doc-convert'
        }, makePanel(ctx)))
      }
    }
  }
})
