import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import {readFileSync} from 'node:fs'
import {clean,esc} from '../src/core/text.js'
const source=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
test('migration status and matching groups display distinct current-to-VF pairs',()=>{
 const session={references:{itemMasters:{entries:[{name:'VF_EL_PANEL'},{name:'VF_EL_PUMP'}]}}};
 const context=vm.createContext({S:{session},clean,esc,modifySuggestedFix:f=>f.expected?{changes:[{prop:'itemMaster',before:f.actual,value:f.expected}]}:null,modifyPatternKey:(_id,key)=>key});
 vm.runInContext(source.slice(source.indexOf('function modifyItemMasterSwap('),source.indexOf('/* Findings with the same explanation')),context);
 vm.runInContext(source.slice(source.indexOf('function modifyPatterns('),source.indexOf('function modifyPatternCountText(')),context);
 const api=vm.runInContext('({modifyItemMasterSwap,modifyItemMasterCatalogStatus,modifyPatterns})',context),rule={id:'item-master.migration-advisory'};
 assert.match(api.modifyItemMasterCatalogStatus(rule),/VF catalog loaded/);
 const findings=['PANEL','PANEL','PUMP','PUMP'].map(name=>({rule,why:'Migration available',actual:`CAMPUS_EL_${name}`,expected:`VF_EL_${name}`}));
 const groups=api.modifyPatterns({rule,matches:findings}).groups;
 assert.equal(groups.length,2);assert.ok(groups.every(group=>group.findings.length===2));
 const html=api.modifyItemMasterSwap(findings[0]);assert.match(html,/Current:.*CAMPUS_EL_PANEL/);assert.match(html,/Suggested VF:.*VF_EL_PANEL/);
 assert.match(api.modifyItemMasterSwap({...findings[0],expected:null}),/No unique VF replacement/);
 session.references={};assert.match(api.modifyItemMasterCatalogStatus(rule),/No local VF catalog/);
});
