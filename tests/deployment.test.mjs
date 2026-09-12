import { SSM_AUDIT_RULES } from '../src/audit/engine.js'
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
  /* Release notes describe active behavior only. Never the words disabled/inactive,
     and never a feature area whose rules are currently switched off — the banned
     list is derived from the engine, so re-enabling a rule family lifts the ban. */
  const notes=changelog.map(entry=>entry.notes).join('\n')
  assert.doesNotMatch(notes,/\b(?:disabled|inactive)\b/i)
  const disabledCategories=new Set(Object.values(SSM_AUDIT_RULES).filter(rule=>!rule.enabled).map(rule=>rule.category))
  const bannedWords={milestones:/\bmilestones?\b/i,'item-masters':/\bitem masters?\b/i}
  for(const [category,pattern] of Object.entries(bannedWords))if(disabledCategories.has(category))assert.doesNotMatch(notes,pattern,`release notes mention ${category} while its rules are disabled`)
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

test('source and fixtures exclude confidential targets apart from the explicitly approved milestone mapping',()=>{
  const paths=[];const walk=directory=>{for(const entry of readdirSync(directory,{withFileTypes:true})){const path=resolve(directory,entry.name);if(entry.isDirectory()&&entry.name!=='.git')walk(path);else if(entry.isFile()&&entry.name!=='.confidential-terms')paths.push(path);}};walk(root)
  const confidentialNames=new RegExp([['Spar','row'].join(''),['Exto-Cx-Registry','_SP'].join('')].join('|'),'i')
  assert.equal(paths.filter(path=>/\.xlsx$/i.test(path)).map(path=>relative(root,path)).join(','),'tests/fixtures/synthetic-registry.xlsx')
  assert.equal(createHash('sha256').update(readFileSync(resolve(root,'tests/fixtures/synthetic-registry.xlsx'))).digest('hex'),'e6d75eec5f5f8fb20ba8bb8b6d96c4fadd024a46f7adf988f7d7aea951c419da')
  const siteName=['Spar','row'].join(''),factory=`audit${siteName}MilestoneMigration`;
  for(const path of paths.filter(path=>!/\.xlsx$|sheetjs\.js$|SSM-Audit\.html$/.test(path))){
    const name=relative(root,path);let text=readFileSync(path,'utf8');
    // Publication approval is limited to the photo's exact replacement table.
    if(name==='src/audit/milestone-migration.js'){
      const start=text.indexOf(`export function ${factory}`),end=text.indexOf('function auditMigrationCode');
      assert.ok(start>=0&&end>start);
      const approved=text.slice(start,end);
      assert.equal(createHash('sha256').update(approved).digest('hex'),'10ddd0111dd4855007074b7edecd162cf31f553616d64dcd05175c5a17c4ebbd');
      text=text.slice(0,start)+text.slice(end);
    }
    if(name==='src/ui/guide-content.js')text=text.replace(`built-in ${siteName} L1 replacements`,'built-in replacements').replace(`complete ${siteName} L1 codes`,'complete project L1 codes');
    text=text.replaceAll(factory,'approvedMilestoneMapping');
    assert.equal(confidentialNames.test(text),false,`Unapproved target reference in ${name}`);
  }
})
