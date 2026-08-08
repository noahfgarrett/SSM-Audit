import test from 'node:test'
import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const standalone=resolve(new URL('..',import.meta.url).pathname)
const integrated=resolve(standalone,'../SSManagement')
const shared=['src/audit/model.js','src/audit/engine.js','src/exto/rev21-contract.js']

test('shared audit rules match the local SSManagement source checkout',{skip:!existsSync(integrated)},()=>{
  for(const path of shared)assert.equal(readFileSync(resolve(standalone,path),'utf8'),readFileSync(resolve(integrated,path),'utf8'),`${path} diverged from SSManagement`)
})
