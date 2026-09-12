import test from 'node:test'
import assert from 'node:assert/strict'
import { confidentialHits, confidentialTerms } from '../build/confidential-check.mjs'

test('no tracked text file contains a confidential term', { skip: !confidentialTerms().length && 'no .confidential-terms file' }, () => {
  assert.deepEqual(confidentialHits().map(hit => `${hit.file}:${hit.line}`), [])
})
