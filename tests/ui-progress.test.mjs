import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import {readFileSync} from 'node:fs'

test('a new progress operation cancels the previous delayed hide and exit animation',()=>{
  const source=readFileSync(new URL('../src/ui/feedback.js',import.meta.url),'utf8').replace(/^import .*$/gm,'').replace(/^export /gm,'');
  const nodes=new Map(),timers=new Map();let next=0;
  const node=selector=>{
    if(!nodes.has(selector)){
      const classes=new Set();
      nodes.set(selector,{hidden:true,style:{},setAttribute(){},addEventListener(){},removeEventListener(){},classList:{add:name=>classes.add(name),remove:name=>classes.delete(name),contains:name=>classes.has(name)}});
    }
    return nodes.get(selector);
  };
  const context=vm.createContext({$:node,setTimeout:fn=>{timers.set(++next,fn);return next;},clearTimeout:id=>timers.delete(id)});
  vm.runInContext(source,context);
  const api=vm.runInContext('({showProgress,hideProgress})',context);
  api.showProgress('First');api.hideProgress();assert.equal(timers.size,1);
  api.showProgress('Second');assert.equal(timers.size,0);assert.equal(node('#overlay').hidden,false);assert.equal(node('#lmsg').textContent,'Second');
  api.hideProgress();const [id,hide]=[...timers][0];timers.delete(id);hide();assert.ok(node('#overlay').classList.contains('is-closing'));
  api.showProgress('Third');assert.equal(timers.size,0);assert.equal(node('#overlay').hidden,false);assert.ok(node('#overlay').classList.contains('show'));assert.ok(!node('#overlay').classList.contains('is-closing'));
});

test('progress overlay stacks above action dialogs and Actions exposes direct updated export',()=>{
  const css=readFileSync(new URL('../src/styles/app.css',import.meta.url),'utf8');
  const z=selector=>Number(css.match(new RegExp(`${selector}\\s*\\{[^}]*z-index:\\s*(\\d+)`))[1]);
  assert.ok(z('#overlay')>z('\\.modal-back'));
  const ui=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
  assert.match(ui,/id="exportUpdatedRegistry"/);assert.match(ui,/\$\('#exportUpdatedRegistry'\)\.onclick=[\s\S]*?exportUpdatedRegistryXlsx\(\)/);
});
