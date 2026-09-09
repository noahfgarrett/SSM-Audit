import test from 'node:test'
import assert from 'node:assert/strict'
import { auditReadMilestoneMigration, auditReadMigrationSettings, auditMigrationValue, auditMigrationReferences, auditMilestoneMigrationRows, auditMigrationImpact } from '../src/audit/milestone-migration.js'
import { auditReviewDocument, auditReadReviewDocument } from '../src/audit/actions.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'

const snapshotOf=rows=>auditSnapshotFromAoa([EXTO_REV21_COLUMNS.map(c=>c.header),...rows.map(row=>EXTO_REV21_COLUMNS.map(c=>row[c.field]||''))],{sheet:'Registry'});

const map=(mappings=[{from:'DEMO-L1-M1-01',to:'DEMO-L1-M1-02',label:'DEMO-L1-M1-02 Approved scope',aliases:['Old scope']}])=>({format:'ssm-audit-milestone-map',version:1,project:'Demonstration',mappings});
test('migration matches complete L1 identities, not numeric fragments or L2 values',()=>{
 const profile=auditReadMilestoneMigration(map());
 assert.equal(auditMigrationValue('DEMO-L1-M1-01 Old label',profile).to,'DEMO-L1-M1-02');
 assert.equal(auditMigrationValue('Old scope',profile).to,'DEMO-L1-M1-02');
 for(const value of ['', 'DEMO-L1-M1-010','DEMO-L1-M1-01_1','DEMO-L2-M1-01','OTHER-L1-M1-01','Text mentioning DEMO-L1-M1-01'])assert.equal(auditMigrationValue(value,profile),null);
 assert.equal(auditMigrationValue('DEMO-L1-M1-02 Approved scope',profile),null);
});
test('conflicting, chained, blank and wrong-level mappings fail closed',()=>{
 for(const mappings of [
   [{from:'',to:'DEMO-L1-M1-02',label:'DEMO-L1-M1-02 Scope'}],
   [{from:'DEMO-L2-M1-01',to:'DEMO-L1-M1-02',label:'DEMO-L1-M1-02 Scope'}],
   [...map().mappings,...map().mappings],
   [...map().mappings,{from:'DEMO-L1-M1-02',to:'DEMO-L1-M1-03',label:'DEMO-L1-M1-03 Scope'}],
   [{from:'DEMO-L1-M1-01',to:'DEMO-L1-M1-02',label:'Wrong label'}],
 ])assert.throws(()=>auditReadMilestoneMigration(map(mappings)));
 assert.throws(()=>auditReadMigrationSettings({enabled:true,profile:null}));
});
test('replacements require an enabled local profile and leave blanks alone',()=>{
 const profile=auditReadMilestoneMigration(map()),snapshot=snapshotOf([{equipmentId:'DEMO-1',milestoneParent:'DEMO-L1-M1-01 Old label'},{equipmentId:'DEMO-2'}]);
 assert.equal(auditMilestoneMigrationRows(snapshot,{enabled:false,profile}).length,0);
 assert.equal(auditMilestoneMigrationRows(snapshot,{enabled:true,profile}).length,1);
 assert.equal(snapshot.rows[0].milestoneParent,'DEMO-L1-M1-01 Old label');
});
test('expected parent mappings overlay references without mutating the uploaded register',()=>{
 const profile=auditReadMilestoneMigration(map()),references={milestones:{kind:'milestones',entries:[{id:'DEMO-L2-M1-10',parentId:'DEMO-L1-M1-01',parentTitle:'Old scope'}]}};
 const result=auditMigrationReferences(references,{enabled:true,profile});
 assert.equal(result.milestones.entries[0].parentId,'DEMO-L1-M1-02');
 assert.equal(references.milestones.entries[0].parentId,'DEMO-L1-M1-01');
 assert.equal(auditMigrationReferences(references,{enabled:false,profile}),references);
});
test('saved reviews preserve the project mapping and opt-in state; old reviews still load',async()=>{
 const profile=auditReadMilestoneMigration(map()),snapshot=snapshotOf([{equipmentId:'DEMO-1',milestoneParent:'DEMO-L1-M1-01 Old label'}]);
 const settings={enabled:true,profile},doc=await auditReviewDocument({baselineSnapshot:snapshot,milestoneMigration:settings});
 assert.deepEqual((await auditReadReviewDocument(doc,snapshot)).milestoneMigration,settings);
 delete doc.milestoneMigration;assert.deepEqual((await auditReadReviewDocument(doc,snapshot)).milestoneMigration,{enabled:false,profile:null});
});
test('explicit L1 migrations retain pairing flags but never waive unrelated errors or arbitrary L1 edits',()=>{
 const profile=auditReadMilestoneMigration(map()),settings={enabled:true,profile};
 const before=snapshotOf([{equipmentId:'DEMO-1',milestoneParent:'DEMO-L1-M1-01 Old label',milestone:'DEMO-L2-M1-10'}]);
 const after=snapshotOf([{equipmentId:'DEMO-1',milestoneParent:profile.mappings[0].label,milestone:'DEMO-L2-M1-10'}]);
 const flag={rule:{id:'reference.milestone-parent-mismatch'},sheet:'Registry',row:before.rows[0]._source.row,severity:'error'};
 const other={...flag,rule:{id:'parent.cross-upn'}},impact={unsafe:[flag,other],introduced:[flag],changed:[],resolved:[]};
 const result=auditMigrationImpact(impact,before,after,settings);
 assert.deepEqual(result.unsafe,[other]);assert.deepEqual(result.migrationConflicts,[flag]);assert.deepEqual(result.introduced,[flag]);
 assert.equal(auditMigrationImpact(impact,before,after,{enabled:false,profile}),impact);
 const wrong=snapshotOf([{equipmentId:'DEMO-1',milestoneParent:'DEMO-L1-M1-99 Other',milestone:'DEMO-L2-M1-10'}]);
 assert.deepEqual(auditMigrationImpact(impact,before,wrong,settings).unsafe,[flag,other]);
});
