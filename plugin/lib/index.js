/**
 * dsh-deliverable-qa 宿主插件。
 *
 * 两件事:
 *  1. 注册 qa_check 工具:对 md/txt/docx/pptx 跑排版 / AI 腔 / 敏感内容三线质检;
 *  2. 在 tools/pre-execute 瀑布上挂交付前安检门:present 声明交付物之前自动质检,
 *     发现高危问题(密钥、凭证等)时按配置拦截或告警。
 *
 * 只使用 ctx API 与 Node 内建能力,不导入 DSH 运行时包(defineTool 除外)。
 */
import { defineTool } from '@deepseek-ai/dsh-tools'
import { CHECKER_PATH, fingerprint, isCheckable, runChecker, summarize } from './runner.js'

export const name = 'deliverable-qa'
export const inject = ['tools']

const GATES = ['block', 'warn', 'off']

function normalizeConfig(config) {
  const raw = config && typeof config === 'object' ? config : {}
  return {
    gate: GATES.includes(raw.gate) ? raw.gate : 'block',
    python: typeof raw.python === 'string' ? raw.python : '',
    configPath: typeof raw.configPath === 'string' ? raw.configPath : '',
    timeoutMs: Number.isFinite(raw.timeoutMs) ? raw.timeoutMs : 20000,
  }
}

/** 从 present 调用参数中取出文件路径。 */
function presentPaths(args) {
  const files = args && Array.isArray(args.files) ? args.files : []
  return files
    .map((entry) => (entry && typeof entry.path === 'string' ? entry.path : ''))
    .filter(Boolean)
}

function formatIssue(issue) {
  return '- ' + issue.file + ' ' + issue.loc + ' [' + issue.rule + '] ' + issue.message
}

function renderReport(value) {
  const lines = ['# 交付物质检报告', '']
  lines.push('- 检查文件:' + value.files.length + ' 个')
  lines.push('- 问题总数:' + value.total + '(高危 ' + value.high + ')')
  lines.push('')
  if (value.issues.length > 0) {
    lines.push('| 文件 | 位置 | 检查线 | 严重度 | 问题 | 建议 |')
    lines.push('| --- | --- | --- | --- | --- | --- |')
    for (const issue of value.issues) {
      lines.push('| ' + issue.file + ' | ' + issue.loc + ' | ' + issue.line + ' | ' + issue.severity + ' | ' + issue.message + ' | ' + issue.advice + ' |')
    }
  } else {
    lines.push('未发现问题。')
  }
  if (value.high > 0) {
    lines.push('')
    lines.push('存在高危问题:请先处理(删除/轮换凭证),不要直接交付。')
  }
  return lines.join('\n')
}

export function apply(ctx, config) {
  const options = normalizeConfig(config)
  console.log('[deliverable-qa] 已加载:交付门=' + options.gate + ',质检脚本=' + CHECKER_PATH)
  /** path -> { stamp, result } :同一文件未变化时不重复跑 Python。 */
  const cache = new Map()

  async function check(paths) {
    const pending = []
    const cached = []
    for (const file of paths) {
      const stamp = fingerprint(file)
      const hit = stamp ? cache.get(file) : undefined
      if (hit && hit.stamp === stamp) cached.push(hit.result)
      else pending.push({ file, stamp })
    }
    let fresh = null
    if (pending.length > 0) {
      const run = await runChecker(pending.map((item) => item.file), options)
      if (!run.ok) return { ok: false, error: run.error }
      fresh = run.data
      const summary = summarize(fresh)
      for (const item of pending) {
        if (item.stamp) cache.set(item.file, { stamp: item.stamp, result: summary })
      }
    }
    for (const result of cached) {
      // 命中缓存的文件不重复计入:仅用于判断是否存在高危。
      if (result.high.length > 0) return { ok: true, data: fresh, cachedHigh: result.high }
    }
    if (fresh === null) return { ok: true, data: null, cachedHigh: [] }
    return { ok: true, data: fresh, cachedHigh: [] }
  }

  const preExecute = async (exec, next) => {
    try {
      if (options.gate === 'off' || !exec || exec.name !== 'present') return next()
      const targets = presentPaths(exec.arguments).filter(isCheckable)
      if (targets.length === 0) return next()
      const run = await check(targets)
      if (!run.ok) {
        console.warn('[deliverable-qa] 交付前质检未完成,已放行:' + run.error)
        return next()
      }
      const freshHigh = run.data ? summarize(run.data).high : []
      const high = [...run.cachedHigh || [], ...freshHigh]
      if (high.length === 0) return next()
      const detail = high.map(formatIssue).join('\n')
      if (options.gate === 'block') {
        return {
          kind: 'deny',
          reason: '交付前质检发现 ' + high.length + ' 处高危问题(密钥/凭证等),已拦截本次交付。请先处理再交付:\n' + detail,
        }
      }
      console.warn('[deliverable-qa] 交付物存在 ' + high.length + ' 处高危问题:\n' + detail)
    } catch (error) {
      console.warn('[deliverable-qa] 交付前质检异常,已放行:' + String((error && error.message) || error))
    }
    return next()
  }

  ctx.on('tools/pre-execute', preExecute, { global: true })

  ctx.tools.register(defineTool({
    name: 'qa_check',
    description: '对交付文档(md/txt/docx/pptx)做质检:排版规范、AI 腔、敏感内容三线检查,返回问题清单与修复建议。文档即将交付前使用;present 声明交付物时也会自动质检。',
    parameters: {
      paths: {
        type: 'array',
        required: true,
        description: '待质检文件路径(绝对路径或会话工作目录相对路径)',
        items: { type: 'string' },
      },
      only: {
        type: 'string',
        enum: ['typography', 'aitone', 'sensitive'],
        description: '只跑指定检查线:typography 排版 / aitone AI腔 / sensitive 敏感内容',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        required: true,
        properties: {
          files: { type: 'array', required: true, items: { type: 'string' } },
          total: { type: 'integer', required: true },
          high: { type: 'integer', required: true },
          issues: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              required: true,
              properties: {
                file: { type: 'string', required: true },
                loc: { type: 'string', required: true },
                line: { type: 'string', required: true },
                rule: { type: 'string', required: true },
                severity: { type: 'string', required: true },
                message: { type: 'string', required: true },
                advice: { type: 'string', required: true },
              },
            },
          },
        },
      },
      render: (_args, value) => [{ type: 'text', text: renderReport(value) }],
    },
    async execute(args) {
      const run = await runChecker(args.paths || [], { ...options, only: args.only })
      if (!run.ok) throw new Error(run.error)
      const summary = summarize(run.data)
      return {
        files: (run.data.files || []).map((entry) => entry.path),
        total: summary.total,
        high: summary.high.length,
        issues: summary.issues.map((issue) => ({
          file: issue.file,
          loc: issue.loc,
          line: issue.line,
          rule: issue.rule,
          severity: issue.severity,
          message: issue.message,
          advice: issue.advice,
        })),
      }
    },
  }))
}
