import test from 'node:test'
import assert from 'node:assert/strict'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { duplicateTopLevelNames } from '../build/unique-names.mjs'
import { MODULES } from '../build/manifest.mjs'

test('no two bundled modules declare the same top-level name', () => {
  assert.deepEqual(duplicateTopLevelNames(fileURLToPath(new URL('..', import.meta.url)), MODULES), [])
})

test('a repeated top-level name across modules is reported', () => {
  const dir = mkdtempSync(join(tmpdir(), 'ssm-names-'))
  writeFileSync(join(dir, 'a.js'), "export function shared(){}\nconst onlyA=1\n")
  writeFileSync(join(dir, 'b.js'), "function inner(){ function shared(){} }\nexport const shared=2\n")
  assert.deepEqual(duplicateTopLevelNames(dir, ['a.js', 'b.js']), ['shared (a.js and b.js)'])
})
