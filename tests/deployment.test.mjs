import test from 'node:test'
import assert from 'node:assert/strict'
import { execFileSync } from 'node:child_process'
import { readFileSync, readdirSync } from 'node:fs'
import { resolve } from 'node:path'

const root=resolve(new URL('..',import.meta.url).pathname)
test('single-file build is offline-ready and pinned to its own updater',()=>{
  execFileSync(process.execPath,['build/build.mjs'],{cwd:root,stdio:'pipe'})
  const html=readFileSync(resolve(root,'SSM-Audit.html'),'utf8')
  assert.match(html,/<title>SSM Audit<\/title>/)
  assert.match(html,/noahfgarrett\/SSM-Audit-Releases/)
  assert.doesNotMatch(html,/SSManagement-Releases|SSM-Builder-Releases/i)
  assert.doesNotMatch(html,/src=["']https?:|href=["']https?:/i)
  assert.equal((html.match(/api\.github\.com\/repos\//g)||[]).length,1)
  assert.match(html,/SSM-Audit-v\$\{String\(version\)/)
})

test('committed source and fixtures contain no confidential audit target names',()=>{
  const paths=[];const walk=directory=>{for(const entry of readdirSync(directory,{withFileTypes:true})){const path=resolve(directory,entry.name);if(entry.isDirectory()&&entry.name!=='.git')walk(path);else if(entry.isFile())paths.push(path);}};walk(root)
  const text=paths.filter(path=>!/\.xlsx$|sheetjs\.js$|SSM-Audit\.html$/.test(path)).map(path=>readFileSync(path,'utf8')).join('\n')
  const confidentialNames=new RegExp([['Spar','row'].join(''),['Exto-Cx-Registry','_SP'].join('')].join('|'),'i')
  assert.doesNotMatch(text,confidentialNames)
})
