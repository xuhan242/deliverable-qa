/**
 * dsh-deliverable-qa 运行时支撑:定位 Python、调用质检脚本、解析 JSON 报告。
 * 只使用 Node 内建能力,不导入任何 DSH 运行时包。
 */
import { spawn } from 'node:child_process'
import { existsSync, readdirSync, statSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, extname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const PKG_ROOT = dirname(dirname(fileURLToPath(import.meta.url)))
export const CHECKER_PATH = join(PKG_ROOT, 'checker', 'qa_check.py')
const CHECKABLE = new Set(['.md', '.markdown', '.txt', '.docx', '.pptx'])

/** 该扩展名是否属于可质检的文档类型。 */
export function isCheckable(file) {
  return CHECKABLE.has(extname(String(file || '')).toLowerCase())
}

function dshHome() {
  const explicit = process.env.DSH_HOME
  if (explicit && explicit.trim()) return explicit.trim()
  return join(homedir(), '.dsh')
}

/** DSH 自带运行时随包分发的 Python(桌面版),按存在性收集。 */
function bundledPythons() {
  const found = []
  const root = join(dshHome(), 'dsh-runtimes')
  if (!existsSync(root)) return found
  for (const id of readdirSync(root)) {
    const base = join(root, id, 'dependencies', 'python')
    for (const rel of ['python.exe', 'bin/python3', 'bin/python', 'python3', 'python']) {
      const candidate = join(base, rel)
      if (existsSync(candidate)) found.push(candidate)
    }
  }
  return found
}

/** Python 候选顺序:显式配置 → DSH 自带运行时 → PATH。 */
export function pythonCandidates(configured) {
  const list = []
  if (configured && String(configured).trim()) list.push(String(configured).trim())
  list.push(...bundledPythons())
  list.push(process.platform === 'win32' ? 'python.exe' : 'python3', 'python')
  return [...new Set(list)]
}

function spawnOnce(python, args, timeoutMs) {
  return new Promise((resolve) => {
    let child
    try {
      child = spawn(python, args, {
        windowsHide: true,
        env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
      })
    } catch (error) {
      resolve({ failed: true, missing: true, error: String((error && error.message) || error) })
      return
    }
    let stdout = ''
    let stderr = ''
    let settled = false
    const finish = (value) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolve(value)
    }
    const timer = setTimeout(() => {
      try { child.kill() } catch {}
      finish({ failed: true, error: '质检超时(' + timeoutMs + 'ms)' })
    }, timeoutMs)
    child.stdout.on('data', (chunk) => { stdout += chunk })
    child.stderr.on('data', (chunk) => { stderr += chunk })
    child.on('error', (error) => {
      finish({ failed: true, missing: !!(error && error.code === 'ENOENT'), error: String((error && error.message) || error) })
    })
    child.on('close', (code) => finish({ failed: false, code, stdout, stderr }))
  })
}

/** 运行质检,逐个尝试 Python 候选,返回解析后的报告。 */
export async function runChecker(files, options = {}) {
  const targets = (files || []).map(String).filter(isCheckable)
  if (targets.length === 0) return { ok: false, error: '没有可质检的文档(md/txt/docx/pptx)' }
  if (!existsSync(CHECKER_PATH)) return { ok: false, error: '找不到质检脚本:' + CHECKER_PATH }
  const args = [CHECKER_PATH]
  if (options.only) args.push('--only', String(options.only))
  if (options.configPath) args.push('--config', String(options.configPath))
  args.push('--json', ...targets)
  const timeoutMs = Number.isFinite(options.timeoutMs) ? options.timeoutMs : 20000
  let lastError = '未找到可用的 Python 解释器'
  for (const python of pythonCandidates(options.python)) {
    const run = await spawnOnce(python, args, timeoutMs)
    if (run.failed) {
      lastError = run.error
      if (run.missing) continue
      return { ok: false, error: run.error }
    }
    if (run.code === 2) return { ok: false, error: (run.stderr || '').trim() || '质检脚本参数或配置错误' }
    try {
      return { ok: true, data: JSON.parse(run.stdout), python }
    } catch (error) {
      const detail = run.stderr ? ' / ' + run.stderr.trim().slice(0, 300) : ''
      return { ok: false, error: '质检输出无法解析:' + String((error && error.message) || error) + detail }
    }
  }
  return { ok: false, error: lastError }
}

/** 统计报告:总数与高危条目。 */
export function summarize(data) {
  const issues = Array.isArray(data && data.issues) ? data.issues : []
  const high = issues.filter((issue) => issue && issue.severity === '高')
  return { total: issues.length, high, issues }
}

/** 文件当前指纹,用于跳过重复质检。 */
export function fingerprint(file) {
  try {
    const stat = statSync(file)
    return stat.mtimeMs + ':' + stat.size
  } catch {
    return ''
  }
}
