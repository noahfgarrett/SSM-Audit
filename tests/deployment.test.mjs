import test from 'node:test'
import assert from 'node:assert/strict'
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { readFileSync, readdirSync } from 'node:fs'
import { relative, resolve } from 'node:path'

const root=resolve(new URL('..',import.meta.url).pathname)
test('release versions stay synchronized and changelog describes active behavior only',()=>{
  const packageJson=JSON.parse(readFileSync(resolve(root,'package.json'),'utf8'))
  const versionSource=readFileSync(resolve(root,'src/version.js'),'utf8')
  const changelog=JSON.parse(readFileSync(resolve(root,'src/changelog.json'),'utf8'))
  assert.match(versionSource,new RegExp(`APP_VERSION=['\"]${packageJson.version.replace(/\./g,'\\.')}['\"]`))
  assert.equal(changelog[0].version,packageJson.version)
  assert.doesNotMatch(changelog.map(entry=>entry.notes).join('\n'),/\b(?:disabled|inactive|milestones?|item masters?)\b/i)
})

test('single-file build is offline-ready and pinned to its own updater',()=>{
  execFileSync(process.execPath,['build/build.mjs'],{cwd:root,stdio:'pipe'})
  const html=readFileSync(resolve(root,'SSM-Audit.html'),'utf8')
  const appScript=html.match(/<script>\s*"use strict";([\s\S]*?)<\/script>/)
  assert.ok(appScript,'built application script is present')
  execFileSync(process.execPath,['--check','-'],{input:appScript[1],stdio:'pipe'})
  assert.match(html,/<title>SSM Audit<\/title>/)
  assert.match(html,/SSM hierarchy/)
  assert.match(html,/noahfgarrett\/SSM-Audit-Releases/)
  assert.doesNotMatch(html,/SSManagement-Releases|SSM-Builder-Releases/i)
  assert.doesNotMatch(html,/src=["']https?:|href=["']https?:/i)
  assert.equal((html.match(/api\.github\.com\/repos\//g)||[]).length,1)
  assert.match(html,/SSM-Audit-v\$\{String\(version\)/)
})

test('committed source and fixtures contain no confidential audit target names',()=>{
  const paths=[];const walk=directory=>{for(const entry of readdirSync(directory,{withFileTypes:true})){const path=resolve(directory,entry.name);if(entry.isDirectory()&&entry.name!=='.git')walk(path);else if(entry.isFile())paths.push(path);}};walk(root)
  const confidentialNames=new RegExp([['Spar','row'].join(''),['Exto-Cx-Registry','_SP'].join('')].join('|'),'i')
  assert.equal(paths.filter(path=>/\.xlsx$/i.test(path)).map(path=>relative(root,path)).join(','),'tests/fixtures/synthetic-registry.xlsx')
  assert.equal(createHash('sha256').update(readFileSync(resolve(root,'tests/fixtures/synthetic-registry.xlsx'))).digest('hex'),'e6d75eec5f5f8fb20ba8bb8b6d96c4fadd024a46f7adf988f7d7aea951c419da')
  const text=paths.filter(path=>!/\.xlsx$|sheetjs\.js$|SSM-Audit\.html$/.test(path)).map(path=>readFileSync(path,'utf8')).join('\n')
  assert.doesNotMatch(text,confidentialNames)
})
