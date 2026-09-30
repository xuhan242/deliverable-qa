/**
 * 把仓库根目录的质检引擎同步进插件包:qa_check.py 与 rules/ 是唯一来源。
 * 用法:npm run sync(在 plugin/ 目录)
 */
import { cpSync, existsSync, mkdirSync, rmSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const pluginRoot = dirname(here)
const repoRoot = dirname(pluginRoot)
const source = join(repoRoot, 'qa_check.py')
const rules = join(repoRoot, 'rules')
const target = join(pluginRoot, 'checker')

if (!existsSync(source) || !existsSync(rules)) {
  console.error('未找到仓库根目录的 qa_check.py / rules/,请在完整仓库中运行')
  process.exit(1)
}
rmSync(target, { recursive: true, force: true })
mkdirSync(target, { recursive: true })
cpSync(source, join(target, 'qa_check.py'))
cpSync(rules, join(target, 'rules'), { recursive: true })
console.log('已同步质检引擎到 ' + target)
