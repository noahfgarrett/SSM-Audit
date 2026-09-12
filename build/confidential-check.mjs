/* Fails when any term listed in the git-ignored .confidential-terms file
   appears in tracked (or, with --staged, staged) text files. One term per
   line, matched case-insensitively; blank lines and # comments are ignored.
   Output names file and line only, never the term itself. */
import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(fileURLToPath(new URL('..', import.meta.url)))
const TEXT_FILE = /\.(js|mjs|cjs|json|html|css|md|txt|csv|yml|yaml|xml|svg|sh)$/i

export function confidentialTerms() {
  const path = resolve(root, '.confidential-terms')
  if (!existsSync(path)) return []
  return readFileSync(path, 'utf8').split(/\r?\n/).map(line => line.trim()).filter(line => line && !line.startsWith('#'))
}
function git(args) { return execFileSync('git', args, { cwd: root, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }) }
export function confidentialHits({ staged = false } = {}) {
  const terms = confidentialTerms().map(term => term.toLowerCase())
  if (!terms.length) return []
  const files = git(staged ? ['diff', '--cached', '--name-only', '--diff-filter=ACMR'] : ['ls-files']).split('\n').filter(file => file && TEXT_FILE.test(file))
  const hits = []
  for (const file of files) {
    let text
    try { text = staged ? git(['show', `:${file}`]) : readFileSync(resolve(root, file), 'utf8') } catch { continue }
    text.split('\n').forEach((line, index) => {
      const lower = line.toLowerCase()
      for (const term of terms) if (lower.includes(term)) { hits.push({ file, line: index + 1 }); break }
    })
  }
  return hits
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const hits = confidentialHits({ staged: process.argv.includes('--staged') })
  for (const hit of hits) console.error(`${hit.file}:${hit.line} contains a confidential term`)
  process.exit(hits.length ? 1 : 0)
}
