import { readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { MODULES } from './manifest.mjs'

const root=resolve(fileURLToPath(new URL('..',import.meta.url)));
const read=path=>readFileSync(resolve(root,path),'utf8');
function classicScript(source,path){
  const output=source
    .replace(/^\s*import\s+[^\n]+\n/gm,'')
    .replace(/^\s*export\s+\{[^}]*\};?\s*$/gm,'')
    .replace(/\bexport\s+(?=(?:async\s+)?function\b|class\b|const\b|let\b|var\b)/g,'');
  if(/^\s*(?:import|export)\b/m.test(output))throw new Error(`ES module syntax survived in ${path}`);
  return `\n/* ${path} */\n${output.trim()}\n`;
}
function replaceOnce(source,marker,value){const count=source.split(marker).length-1;if(count!==1)throw new Error(`Expected one ${marker}, found ${count}`);return source.replace(marker,()=>value);}

const app=MODULES.map(path=>classicScript(read(path),path)).join('');
const changelog=JSON.parse(read('src/changelog.json'));
let html=read('src/index.html');
html=replaceOnce(html,'<!--@inject:styles-->',read('src/styles/app.css'));
html=replaceOnce(html,'<!--@inject:vendor:sheetjs-->',read('src/vendor/sheetjs.js'));
html=replaceOnce(html,'<!--@inject:changelog-->',`const CHANGELOG=Object.freeze(${JSON.stringify(changelog)});`);
html=replaceOnce(html,'<!--@inject:app-->',app);
if(/@inject:/.test(html))throw new Error('A build injection marker remains');
if(!html.includes("UPDATE_REPOSITORY='noahfgarrett/SSM-Audit-Releases'"))throw new Error('Standalone updater repository is not pinned');
if(/SSManagement-Releases|SSM-Builder/i.test(app))throw new Error('A foreign product update channel leaked into the standalone build');
writeFileSync(resolve(root,'SSM-Audit.html'),html);
console.log(`Built SSM-Audit.html (${Buffer.byteLength(html).toLocaleString()} bytes)`);
