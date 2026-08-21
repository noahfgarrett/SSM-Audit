import { EXTO_REV21_COLUMNS, extoRev21SystemsForUpn } from '../exto/rev21-contract.js'

/* ---- Worked examples for the Rules screen ----
   One small, made-up registry per check, built so the check fires on exactly the
   row it is meant to illustrate. `focus` names the cells to shade; `marks`
   optionally splits one cell into plain / offending runs so the exact characters
   stand out. Tags, buildings and descriptions are invented ("B1-…"); UPNs,
   System Names, disciplines and VF Item Master names are the VF Exto Upload Template / VF
   vocabulary that already ships in the app. tests/examples.test.mjs runs the
   real engine over every example and asserts it flags the focused row. */

const sys=upn=>(extoRev21SystemsForUpn(upn)||[])[0]||'';
const MECH='MECHANICAL WET',MECH_EXH='MECHANICAL EXHAUST',ELEC='ELECTRICAL',FMS='FACILITIES MONITORING SYSTEM',LSS='LIFE SAFETY SYSTEM',UPW='UPW';
const B='B1';

/* r(...) writes one registry row in the VF Exto Upload Template field names the engine reads. */
function r(values){
  return {
    equipmentId:values.id||'',equipmentDescription:values.desc||'',closestParent:values.parent||'',
    /* A row whose parent is its own System Name is a system root; registries mark those New. */
    closestParentStatus:values.status||(values.parent&&values.parent===(values.sys!==undefined?values.sys:(values.upn?sys(values.upn):''))?'NEW':''),
    dependencies:values.deps||'',dependencyProject:values.depProject||'',upn:values.upn||'',systemName:values.sys!==undefined?values.sys:(values.upn?sys(values.upn):''),
    discipline:values.disc||'',building:values.bldg||B,itemMaster:values.im||'',milestone:values.l2||'',milestoneParent:values.l1||'',
    equipmentClassification:values.cls||'',
  };
}
function ex(rule,example){return [rule,Object.freeze({rows:Object.freeze(example.rows.map(Object.freeze)),focus:Object.freeze(example.focus||[]),marks:Object.freeze(example.marks||[]),caption:example.caption||'',fix:example.fix||'',options:example.options||null,columns:example.columns||null})];}

/* A routine mechanical system the electrical/controls examples can lean on. */
const AHU={id:'B1-AHU-1041',desc:'Air handler',parent:sys('104'),upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'};
const PANEL={id:'B1-PNL-6041',desc:'Electrical panel',parent:sys('604'),upn:'604',disc:ELEC,im:'VF_EL_PANEL'};
const PLC={id:'B1-PLC-6501',desc:'PLC controller',parent:sys('650'),upn:'650',disc:FMS,im:'VF_I&C_PLC_TELE'};

export const SSM_AUDIT_EXAMPLES=Object.freeze(Object.fromEntries([
  /* ---------------- registry integrity ---------------- */
  ex('identity.duplicate-equipment-id',{
    rows:[r({...AHU}),r({id:'B1-AHU-1041',desc:'Air handler (duplicate row)',parent:sys('104'),upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'}),r({id:'B1-AHU-1042',desc:'Air handler',parent:sys('104'),upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:0,field:'equipmentId'},{row:1,field:'equipmentId'}],
    caption:'B1-AHU-1041 is on two rows. Only one of them can be the real equipment.',fix:'Delete the duplicate, or give the second unit its own tag (B1-AHU-1043).'}),
  ex('parent.blank',{
    rows:[r({id:'B1-AHU-1041',desc:'Air handler',parent:'',upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:0,field:'closestParent'}],
    caption:'The Closest Parent is empty, so Exto cannot place this row anywhere in the SSM.',fix:`Name the parent equipment — or, when this is the top of the system, its System Name: ${sys('104')}.`}),
  ex('parent.self',{
    rows:[r({id:'B1-AHU-1041',desc:'Air handler',parent:'B1-AHU-1041',upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:0,field:'closestParent'}],
    caption:'The row names itself as its own parent — a loop one row long.',fix:`Point Closest Parent at the real parent, or the System Name ${sys('104')}.`}),
  ex('parent.unresolved',{
    rows:[r({...AHU}),r({id:'B1-SF-1041',desc:'Supply fan',parent:'B1-AHU-1O41',upn:'104',disc:MECH,im:'VF_MD_EXHAUST FAN'})],
    focus:[{row:1,field:'closestParent'}],
    marks:[{row:1,field:'closestParent',parts:[['B1-AHU-1',false],['O',true],['41',false]]}],
    caption:'The parent is typed with the letter O instead of the digit 0, so no row in the registry matches it.',fix:'Correct the tag to B1-AHU-1041 (or set Closest Parent Status to Existing if it really lives in another project).'}),
  ex('parent.generated-header-review',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:'CHW PUMP GROUP',status:'NEW',upn:'111',disc:MECH,im:'VF_Rotating_PUMP'}),r({id:'B1-PMP-1112',desc:'Chilled water pump',parent:'CHW PUMP GROUP',status:'NEW',upn:'111',disc:MECH,im:'VF_Rotating_PUMP'})],
    focus:[{row:0,field:'closestParent'},{row:1,field:'closestParent'}],
    caption:'"CHW PUMP GROUP" is marked New and is not any equipment row — it would be created as a header on upload.',fix:'If the header is intended, add a row for it with a VF_Blank Item Master; otherwise parent the pumps to equipment or the System Name.'}),
  ex('parent.cross-upn',{
    rows:[r({...PANEL}),r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:'B1-PNL-6041',deps:'',upn:'111',disc:MECH,im:'VF_Rotating_PUMP'})],
    focus:[{row:1,field:'closestParent'},{row:1,field:'upn'}],
    caption:'The pump is on UPN 111 but its parent is on UPN 604. A structural parent must be inside the same UPN.',fix:`Keep B1-PNL-6041 as a dependency and parent the pump inside 111 — to equipment or to ${sys('111')}.`}),
  ex('parent.cross-discipline',{
    rows:[r({id:'B1-XFMR-6021',desc:'Medium voltage transformer',parent:sys('602'),upn:'602',disc:ELEC,im:'VF_EL_MV_XFMR'}),r({id:'B1-PMP-6021',desc:'Cooling pump',parent:'B1-XFMR-6021',upn:'602',disc:MECH,im:'VF_Rotating_PUMP'})],
    focus:[{row:1,field:'closestParent'},{row:1,field:'discipline'}],
    caption:'A mechanical pump is nested under electrical gear. Only controls devices may cross disciplines to sit under the equipment they serve.',fix:'Record the transformer as a dependency and parent the pump inside its own discipline.'}),
  ex('parent.cycle',{
    rows:[r({id:'B1-AHU-1041',desc:'Air handler',parent:'B1-SF-1041',upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'}),r({id:'B1-SF-1041',desc:'Supply fan',parent:'B1-AHU-1041',upn:'104',disc:MECH,im:'VF_MD_EXHAUST FAN'})],
    focus:[{row:0,field:'closestParent'},{row:1,field:'closestParent'}],
    caption:'The AHU’s parent is the fan and the fan’s parent is the AHU — following parents upward never reaches a top.',fix:`Parent the AHU to ${sys('104')} and leave the fan under the AHU.`}),
  ex('dependency.self',{
    rows:[r({id:'B1-AHU-1041',desc:'Air handler',parent:sys('104'),deps:'B1-AHU-1041',upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:0,field:'dependencies'}],
    caption:'The row lists itself as something it depends on.',fix:'Remove the self-reference from Dependencies.'}),
  ex('dependency.duplicate',{
    rows:[r({...PANEL}),r({id:'B1-AHU-1041',desc:'Air handler',parent:sys('104'),deps:'B1-PNL-6041; B1-PNL-6041',upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:1,field:'dependencies'}],
    marks:[{row:1,field:'dependencies',parts:[['B1-PNL-6041; ',false],['B1-PNL-6041',true]]}],
    caption:'The same panel is listed twice on one row.',fix:'Keep one entry per dependency: B1-PNL-6041.'}),
  ex('dependency.unresolved',{
    rows:[r({...PANEL}),r({id:'B1-AHU-1041',desc:'Air handler',parent:sys('104'),deps:'B1-PNL-6O41',upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:1,field:'dependencies'}],
    marks:[{row:1,field:'dependencies',parts:[['B1-PNL-6',false],['O',true],['41',false]]}],
    caption:'The dependency is typed with the letter O, so it matches no tag in the registry and no other project is named.',fix:'Correct it to B1-PNL-6041, or fill Dependency Project when it lives elsewhere.'}),
  ex('dependency.precedence-cycle',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),deps:'B1-PMP-1112',upn:'111',disc:MECH,im:'VF_Rotating_PUMP'}),r({id:'B1-PMP-1112',desc:'Chilled water pump',parent:sys('111'),deps:'B1-PMP-1111',upn:'111',disc:MECH,im:'VF_Rotating_PUMP'})],
    focus:[{row:0,field:'dependencies'},{row:1,field:'dependencies'}],
    caption:'Pump 1111 depends on 1112 and 1112 depends on 1111 — neither can start first.',fix:'Keep the dependency in one direction only (whichever truly must run first).'}),
  ex('dependency.same-upn-bottom-up',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP'}),r({id:'B1-PMP-1112',desc:'Chilled water pump',parent:sys('111'),deps:'B1-PMP-1111',upn:'111',disc:MECH,im:'VF_Rotating_PUMP'})],
    focus:[{row:1,field:'dependencies'}],
    caption:'In a bottom-up (mechanical) discipline, two pumps in the same UPN are normally sequenced by the hierarchy, not by a dependency.',fix:'Confirm the sequencing is intended; otherwise remove the same-system dependency.'}),
  ex('metadata.upn-not-approved',{
    rows:[r({id:'B1-AHU-1041',desc:'Air handler',parent:sys('104'),upn:'1040',sys:sys('104'),disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:0,field:'upn'}],
    marks:[{row:0,field:'upn',parts:[['104',false],['0',true]]}],
    caption:'1040 is not a UPN in the VF Exto Upload Template — UPNs are three digits (or RR / SEC / MISC).',fix:'Use the approved UPN: 104.'}),
  ex('metadata.misc-upn-review',{
    rows:[r({id:'B1-HOIST-0001',desc:'Maintenance hoist',parent:'MISC',sys:'MISC',upn:'MISC',disc:'MECHANICAL MISC',im:'VF_PROC_EQ'})],
    focus:[{row:0,field:'upn'}],
    caption:'The hoist sits on UPN MISC — the catch-all for equipment without a proper system. MISC rows are never wrong by definition, which is exactly why each one deserves a second look.',fix:'Double-check the assignment; if an approved UPN covers this equipment, move the row to it. Keep MISC only when nothing fits.'}),
  ex('metadata.system-upn-mismatch',{
    rows:[r({id:'B1-AHU-1041',desc:'Air handler',parent:sys('111'),upn:'104',sys:sys('111'),disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:0,field:'systemName'},{row:0,field:'upn'}],
    marks:[{row:0,field:'systemName',parts:[['111',true],[sys('111').slice(3),false]]}],
    caption:`The row is on UPN 104 but its System Name belongs to UPN 111.`,fix:`Use the VF Exto Upload Template System Name for 104: ${sys('104')}.`}),
  ex('metadata.ic-discipline',{
    rows:[r({id:'B1-RIO-6501',desc:'Remote I/O panel',parent:sys('650'),upn:'650',disc:'I&C',im:'VF_I&C_RIO W/O SUD'})],
    focus:[{row:0,field:'discipline'}],
    caption:'"I&C" is not the approved controls discipline in the VF Exto Upload Template dropdown.',fix:`Use ${FMS}.`}),
  ex('metadata.upn-inconsistent',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP'}),r({id:'B1-PMP-1112',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH_EXH,im:'VF_Rotating_PUMP'}),r({id:'B1-PMP-1113',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP'})],
    focus:[{row:0,field:'discipline'},{row:1,field:'discipline'}],
    caption:'Rows on UPN 111 carry different disciplines. Every row on a UPN should agree on System Name and Discipline.',fix:`Set all three pumps to ${MECH}.`}),
  ex('identity.tag-looks-like-description',{
    rows:[r({id:'Chilled water pumps',desc:'',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Blank'})],
    focus:[{row:0,field:'equipmentId'}],
    marks:[{row:0,field:'equipmentId',parts:[['Chilled ',true],['water ',true],['pumps',true]]}],
    caption:'Lowercase words in the Equipment ID column read like a description typed where the tag belongs.',fix:'Use a tag (B1-CHW-PUMPS-HDR) and put the words in Equipment Description.'}),
  ex('metadata.classification-not-in-list',{
    rows:[r({id:'B1-XV-1111',desc:'Isolation valve',parent:'B1-PMP-1111',upn:'111',disc:MECH,im:'VF_I&C_VALVE',cls:'XV'}),r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',cls:'PMP'})],
    focus:[{row:0,field:'equipmentClassification'}],
    caption:'"XV" is not in the VF Exto Upload Template Equipment Classification dropdown. It may be a legitimate site code — the count tells you whether to add it.',fix:'Pick the VF Exto Upload Template classification, or raise the site code for addition to the list.'}),

  /* ---------------- milestones ---------------- */
  ex('milestone.incomplete-pair',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',l1:'L1-M1 30% Capacity',l2:'L2-M1-111 30% Capacity'}),r({id:'B1-PMP-1112',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',l1:'',l2:'L2-M1-111 30% Capacity'})],
    focus:[{row:1,field:'milestoneParent'}],
    caption:'This project uses milestones, but the second pump is missing its L1 Milestone Parent.',fix:'Fill both L1 and L2 on every row.'}),
  ex('milestone.l2-upn-mismatch',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',l1:'L1-M1 30% Capacity',l2:'L2-M1-104 30% Capacity'})],
    focus:[{row:0,field:'milestone'},{row:0,field:'upn'}],
    marks:[{row:0,field:'milestone',parts:[['L2-M1-',false],['104',true],[' 30% Capacity',false]]}],
    caption:'The row is on UPN 111 but its L2 milestone names UPN 104.',fix:'Assign the L2 milestone for 111 (L2-M1-111 …), or move the row to the right UPN.'}),
  ex('milestone.intent-mismatch',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',l1:'L1-M1 30% Capacity',l2:'L2-M1-111 100% Capacity'})],
    focus:[{row:0,field:'milestoneParent'},{row:0,field:'milestone'}],
    marks:[{row:0,field:'milestoneParent',parts:[['L1-M1 ',false],['30%',true],[' Capacity',false]]},{row:0,field:'milestone',parts:[['L2-M1-111 ',false],['100%',true],[' Capacity',false]]}],
    caption:'L1 says 30% capacity and L2 says 100% — the two levels describe different phases.',fix:'Make L1 and L2 describe the same phase.'}),
  ex('milestone.parent-inconsistent',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',l1:'L1-M1 30% Capacity',l2:'L2-M1-111 30% Capacity'}),r({id:'B1-PMP-1112',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',l1:'L1-M2 100% Capacity',l2:'L2-M1-111 30% Capacity'})],
    focus:[{row:0,field:'milestoneParent'},{row:1,field:'milestoneParent'}],
    caption:'The same L2 milestone rolls up to a 30% L1 on one row and a 100% L1 on another.',fix:'Give every row of that L2 milestone the same L1 Milestone Parent.'}),
  ex('milestone.local-cohort-outlier',{
    rows:[...Array.from({length:10},(_,index)=>r({id:`B1-PMP-11${String(index+11).padStart(2,'0')}`,desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',l1:'L1-M1 30% Capacity',l2:'L2-M1-111 30% Capacity'})),
      r({id:'B1-PMP-1121',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',l1:'L1-M2 100% Capacity',l2:'L2-M2-111 100% Capacity'})],
    focus:[{row:10,field:'milestoneParent'},{row:10,field:'milestone'}],
    caption:'Eleven identical pumps in the same building and system — ten share one milestone pair, the eleventh does not.',fix:'Confirm the odd pump really belongs to a later phase; otherwise match its peers.'}),
  ex('milestone.level-field-mismatch',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP',l1:'L1-M1 30% Capacity',l2:'L1-M1 30% Capacity'})],
    focus:[{row:0,field:'milestone'}],
    marks:[{row:0,field:'milestone',parts:[['L1',true],['-M1 30% Capacity',false]]}],
    caption:'An L1-shaped identifier is sitting in the L2 Milestone column.',fix:'Put the L2 identifier (L2-M1-111 …) in L2 Milestone.'}),
  ex('milestone.branch-outlier',{
    rows:[r({id:'B1-AHU-1041',desc:'Air handler',parent:sys('104'),upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH',l1:'L1-M1 30% Capacity',l2:'L2-M1-104 30% Capacity'}),
      ...['SF','RF','CC','PF','FF'].map((kind,index)=>r({id:`B1-${kind}-1041`,desc:['Supply fan','Return fan','Cooling coil','Pre-filter','Final filter'][index],parent:'B1-AHU-1041',upn:'104',disc:MECH,im:'VF_MD_EXHAUST FAN',l1:'L1-M1 30% Capacity',l2:'L2-M1-104 30% Capacity'})),
      r({id:'B1-HC-1041',desc:'Heating coil',parent:'B1-AHU-1041',upn:'104',disc:MECH,im:'VF_MD_HEATING COIL',l1:'L1-M2 100% Capacity',l2:'L2-M2-104 100% Capacity'})],
    focus:[{row:6,field:'milestoneParent'},{row:6,field:'milestone'}],
    caption:'The heating coil carries a different milestone pair from its parent AHU and its five sibling components.',fix:'Match the branch unless the later phase is documented.'}),

  /* ---------------- item masters & headers ---------------- */
  ex('item-master.standardized-assignment',{
    rows:[r({id:'B1-AHU-1041',desc:'Air handler',parent:sys('104'),upn:'104',disc:MECH,im:'CA_MD_GAH/GMAH'}),r({id:'B1-AHU-1042',desc:'Air handler',parent:sys('104'),upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'}),r({id:'B1-AHU-1043',desc:'Air handler',parent:sys('104'),upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:0,field:'itemMaster'}],
    marks:[{row:0,field:'itemMaster',parts:[['CA',true],['_MD_GAH/GMAH',false]]}],
    caption:'A legacy project-prefixed Item Master (CA_…) where the VF standard name exists.',fix:'Replace with VF_MD_GAH/GMAH.'}),
  ex('header.item-master-not-blank',{
    rows:[r({id:'B1-CHW-PIPING',desc:'Chilled water piping',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP'}),r({id:'B1-TT-1111',desc:'Temperature transmitter',parent:'B1-CHW-PIPING',upn:'111',disc:FMS,im:'VF_I&C_TRANSMITTER'}),r({id:'B1-TT-1112',desc:'Temperature transmitter',parent:'B1-CHW-PIPING',upn:'111',disc:FMS,im:'VF_I&C_TRANSMITTER'}),r({id:'B1-PT-1111',desc:'Pressure transmitter',parent:'B1-CHW-PIPING',upn:'111',disc:FMS,im:'VF_I&C_TRANSMITTER'})],
    focus:[{row:0,field:'itemMaster'}],
    caption:'Three instruments nest under "Chilled water piping", which only groups them — yet it carries a real pump Item Master, so pump checklists would be applied to piping.',fix:'Give the header VF_Blank.'}),
  ex('header.has-dependency',{
    rows:[r({id:'B1-CHW-PIPING',desc:'Chilled water piping',parent:sys('111'),deps:'B1-PNL-6041',upn:'111',disc:MECH,im:'VF_Blank'}),r({...PANEL}),r({id:'B1-TT-1111',desc:'Temperature transmitter',parent:'B1-CHW-PIPING',upn:'111',disc:FMS,im:'VF_I&C_TRANSMITTER'})],
    focus:[{row:0,field:'dependencies'}],
    caption:'A header (VF_Blank) only groups equipment; it should not depend on anything itself.',fix:'Move the dependency onto the equipment that actually needs the panel.'}),
  ex('header.unused',{
    rows:[r({id:'B1-CHW-PIPING',desc:'Chilled water piping',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Blank'}),r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Rotating_PUMP'})],
    focus:[{row:0,field:'itemMaster'},{row:0,field:'equipmentId'}],
    caption:'A row set up as a header (VF_Blank) with nothing nested under it.',fix:'Nest the equipment it was meant to group under it, or remove the header row.'}),
  ex('dependency.parent-also-listed',{
    rows:[r({id:'B1-SWGR-6021',desc:'Medium voltage switchgear',parent:sys('602'),upn:'602',disc:ELEC,im:'VF_EL_MV_GEAR'}),r({id:'B1-XFMR-6021',desc:'Medium voltage transformer',parent:'B1-SWGR-6021',deps:'B1-SWGR-6021',upn:'602',disc:ELEC,im:'VF_EL_MV_XFMR'})],
    focus:[{row:1,field:'closestParent'},{row:1,field:'dependencies'}],
    caption:'The switchgear is this row’s Closest Parent and is listed as a dependency too. Electrical routinely records the feeder in both places — elsewhere it is worth a look.',fix:'Leave it if deliberate; otherwise remove the duplicate dependency.'}),
  ex('dependency.on-header',{
    rows:[r({id:'B1-CHW-PIPING',desc:'Chilled water piping',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Blank'}),r({id:'B1-TT-1111',desc:'Temperature transmitter',parent:'B1-CHW-PIPING',upn:'111',disc:FMS,im:'VF_I&C_TRANSMITTER'}),r({id:'B1-AHU-1041',desc:'Air handler',parent:sys('104'),deps:'B1-CHW-PIPING',upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:2,field:'dependencies'}],
    caption:'The AHU depends on "Chilled water piping", which is only a header. Dependencies should point at the equipment that is actually needed.',fix:'Depend on the pump or valve inside that header instead.'}),
  ex('parent.cross-building',{
    rows:[r({id:'B1-PNL-6041',desc:'Electrical panel',parent:sys('604'),upn:'604',disc:ELEC,bldg:'B1',im:'VF_EL_PANEL'}),r({id:'B2-PNL-6042',desc:'Electrical panel',parent:'B1-PNL-6041',upn:'604',disc:ELEC,bldg:'B2',im:'VF_EL_PANEL'})],
    focus:[{row:1,field:'closestParent'},{row:1,field:'building'}],
    caption:'The B2 panel is nested under a B1 panel. A structural child stays in its parent’s building.',fix:`Keep B1-PNL-6041 as a dependency and parent the B2 panel inside B2 (or to ${sys('604')}).`}),
  ex('dependency.project-not-needed',{
    rows:[r({...PANEL}),r({id:'B1-AHU-1041',desc:'Air handler',parent:sys('104'),deps:'B1-PNL-6041',depProject:'PH2-EXPANSION',upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'})],
    focus:[{row:1,field:'dependencyProject'}],
    caption:'Dependency Project names another project, but every dependency on the row is a tag in this registry.',fix:'Clear Dependency Project; it is only for dependencies that live elsewhere.'}),
  ex('structure.system-without-root',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:'B1-PMP-1112',upn:'111',disc:MECH,im:'VF_Rotating_PUMP'}),r({id:'B1-PMP-1112',desc:'Chilled water pump',parent:'B1-PMP-1111',upn:'111',disc:MECH,im:'VF_Rotating_PUMP'})],
    focus:[{row:0,field:'systemName'},{row:0,field:'closestParent'},{row:1,field:'closestParent'}],
    caption:`${sys('111')} is in use, but no row sits at the top of it (none has the System Name as its Closest Parent).`,fix:`Make the first pump’s Closest Parent ${sys('111')}.`}),

  /* ---------------- commissioning logic ---------------- */
  ex('logic.control-link-missing',{
    rows:[r({...PLC}),r({id:'B1-TT-1111',desc:'Temperature transmitter',parent:sys('111'),deps:'',upn:'111',disc:FMS,im:'VF_I&C_TRANSMITTER'})],
    focus:[{row:1,field:'closestParent'},{row:1,field:'dependencies'}],
    caption:'A transmitter with no RIO, PLC, VFD or control panel anywhere in its parent or dependencies — nothing controls it.',fix:'Add its RIO or PLC as a dependency (or parent).'}),
  ex('logic.driven-electrical-path-missing',{
    rows:[r({id:'B1-PMP-1111',desc:'Chilled water pump',parent:sys('111'),deps:'',upn:'111',disc:MECH,im:'VF_Rotating_PUMP'})],
    focus:[{row:0,field:'dependencies'}],
    caption:'A pump with no path back to any electrical gear — no panel, MCC, VFD or switchgear as parent or dependency.',fix:'Add the panel or VFD that powers it as a dependency.'}),
  ex('logic.control-electrical-path-missing',{
    rows:[r({id:'B1-RIO-6501',desc:'Remote I/O panel',parent:sys('650'),deps:'',upn:'650',disc:FMS,im:'VF_I&C_RIO W/O SUD'})],
    focus:[{row:0,field:'dependencies'}],
    caption:'An RIO with no path to an electrical supply.',fix:'Add the panel that powers it as a dependency.'}),
  ex('logic.rio-control-path-missing',{
    rows:[r({...PANEL}),r({id:'B1-RIO-6501',desc:'Remote I/O panel',parent:sys('650'),deps:'B1-PNL-6041',upn:'650',disc:FMS,im:'VF_I&C_RIO W/O SUD'})],
    focus:[{row:1,field:'dependencies'}],
    caption:'The RIO has power but names no PLC, I/O cluster or upstream RIO that runs it.',fix:'Add the PLC (B1-PLC-6501) as a dependency.'}),
  ex('logic.drive-parent-unexpected',{
    rows:[r({...PANEL}),r({...AHU}),r({id:'B1-VFD-1041',desc:'Variable frequency drive',parent:'B1-PNL-6041',deps:'B1-PLC-6501',upn:'104',disc:ELEC,im:'VF_EL_VFD'}),r({...PLC})],
    focus:[{row:2,field:'closestParent'}],
    caption:'The VFD is nested under the electrical panel. A drive nests under the equipment it drives; the panel is its power dependency.',fix:'Parent the VFD to B1-AHU-1041 and list B1-PNL-6041 as a dependency.'}),
  ex('sop.vfd-dependencies',{
    rows:[r({...PANEL}),r({...AHU}),r({id:'B1-VFD-1041',desc:'Variable frequency drive',parent:'B1-AHU-1041',deps:'B1-PNL-6041',upn:'104',disc:ELEC,im:'VF_EL_VFD'}),r({...PLC})],
    focus:[{row:2,field:'dependencies'}],
    caption:'Per the SOP a VFD lists both its electrical panel and its PLC. This one has the panel but no PLC.',fix:'Add B1-PLC-6501 to Dependencies.'}),
  ex('logic.heat-trace-chain-missing',{
    rows:[r({id:'B1-XFMR-6101',desc:'Heat trace transformer',parent:sys('610'),upn:'610',disc:ELEC,im:'VF_EL_Branch_XFMR'}),r({id:'B1-HTP-6101',desc:'Heat trace panel',parent:sys('610'),upn:'610',disc:ELEC,im:'VF_EL_HTP'})],
    focus:[{row:1,field:'closestParent'}],
    caption:'The heat-trace panel sits directly under the System Name instead of under its transformer.',fix:'Parent B1-HTP-6101 to B1-XFMR-6101.'}),
  ex('logic.fdu-supported-equipment-missing',{
    rows:[r({id:'B1-FMS-FIBER',desc:'FMS fiber header',parent:sys('650'),upn:'650',disc:FMS,im:'VF_Blank'}),r({...PLC}),r({id:'B1-FDU-6501',desc:'Fiber optic distribution unit (FDU)',parent:'B1-FMS-FIBER',deps:'',upn:'650',disc:FMS,im:'VF_I&C_FDU'}),r({id:'B1-FDU-6502',desc:'Fiber optic distribution unit (FDU)',parent:'B1-PLC-6501',deps:'',upn:'650',disc:FMS,im:'VF_I&C_FDU'})],
    focus:[{row:2,field:'closestParent'},{row:2,field:'dependencies'}],
    caption:'The first FDU hangs off a header with no dependencies, so nothing says what it connects to. The second nests under the PLC — that placement is the relationship, and it needs no dependency.',fix:'Nest B1-FDU-6501 under the PLC, RIO, or patch panel it belongs to (same UPN), or name that equipment as a dependency.'}),
  ex('logic.vesda-fire-alarm-missing',{
    rows:[r({id:'B1-FAP-0001',desc:'Fire alarm panel',parent:sys('630'),upn:'630',disc:LSS,im:'VF_LSS_FAP CARD'}),r({id:'B1-VESDA-0001',desc:'VESDA aspirating smoke detection',parent:sys('630'),deps:'',upn:'630',disc:LSS,im:'VF_LSS_FAP CARD'})],
    focus:[{row:1,field:'dependencies'}],
    caption:'Per the SOP a VESDA system depends on its fire alarm panel. This one lists nothing.',fix:'Add B1-FAP-0001 to Dependencies.'}),

  /* ---------------- SOP nesting ---------------- */
  ex('sop.instrument-parent-upn',{
    rows:[r({...AHU}),r({id:'B1-TT-1111',desc:'Temperature transmitter',parent:'B1-AHU-1041',upn:'104',disc:FMS,im:'VF_I&C_TRANSMITTER'})],
    focus:[{row:1,field:'equipmentId'},{row:1,field:'closestParent'}],
    marks:[{row:1,field:'equipmentId',parts:[['B1-TT-',false],['111',true],['1',false]]}],
    caption:'The transmitter’s tag carries UPN 111, yet it is nested under an AHU in UPN 104 and the row itself says 104.',fix:'Nest it under equipment in 111, or correct the tag if the row is right.'}),
  ex('sop.fms-io-under-vfd',{
    rows:[r({...AHU}),r({id:'B1-VFD-1041',desc:'Variable frequency drive',parent:'B1-AHU-1041',deps:'B1-PNL-6041; B1-PLC-6501',upn:'104',disc:ELEC,im:'VF_EL_VFD'}),r({id:'B1-VFD-1041-IO',desc:'FMS hardwired I/O for VFD',parent:'B1-AHU-1041',deps:'B1-PLC-6501',upn:'104',disc:FMS,im:'VF_I&C_DIGITAL I/O'}),r({...PANEL}),r({...PLC})],
    focus:[{row:2,field:'closestParent'}],
    caption:'The FMS hardwired I/O for the drive is nested under the AHU instead of under the VFD it belongs to.',fix:'Parent B1-VFD-1041-IO to B1-VFD-1041.'}),
  ex('sop.lcp-placement',{
    rows:[r({id:'B1-MAH-1041',desc:'Makeup air handler',parent:sys('104'),upn:'104',disc:MECH,im:'VF_MD_GAH/GMAH'}),r({id:'B1-LCP-1041',desc:'Local control panel',parent:'B1-PNL-6041',upn:'104',disc:FMS,im:'VF_I&C_LCP W/O SUD'}),r({...PANEL})],
    focus:[{row:1,field:'closestParent'}],
    caption:'The local control panel is nested under an electrical panel in another UPN. An LCP sits with the equipment or skid it serves (MAH > SKID > LCP), or is a system root.',fix:'Parent B1-LCP-1041 to B1-MAH-1041 and keep the panel as a dependency.'}),
  ex('sop.untied-instrument-rollup',{
    rows:[r({id:'B1-CHW-PIPING',desc:'Chilled water piping roll-up',parent:sys('111'),upn:'111',disc:MECH,im:'VF_Blank'}),r({id:'B1-PT-1041',desc:'Pressure transmitter',parent:'B1-CHW-PIPING',upn:'111',disc:FMS,im:'VF_I&C_TRANSMITTER'}),r({id:'B1-TT-1111',desc:'Temperature transmitter',parent:'B1-CHW-PIPING',upn:'111',disc:FMS,im:'VF_I&C_TRANSMITTER'}),r({id:'B1-TT-1112',desc:'Temperature transmitter',parent:'B1-CHW-PIPING',upn:'111',disc:FMS,im:'VF_I&C_TRANSMITTER'})],
    focus:[{row:1,field:'equipmentId'},{row:1,field:'closestParent'}],
    marks:[{row:1,field:'equipmentId',parts:[['B1-PT-',false],['104',true],['1',false]]}],
    caption:'An instrument tagged for UPN 104 is parked under a piping roll-up header that belongs to UPN 111.',fix:'Move it under the 104 roll-up or equipment, or fix the tag.'}),
  ex('sop.control-valve-parent',{
    rows:[r({...AHU}),r({id:'B1-TCV-1041',desc:'Temperature control valve',parent:sys('104'),upn:'104',disc:FMS,im:'VF_I&C_VALVE'})],
    focus:[{row:1,field:'closestParent'}],
    caption:'A control valve parented straight to the System Name. Per the SOP it nests under the equipment it serves.',fix:'Parent B1-TCV-1041 to B1-AHU-1041.'}),
  ex('sop.room-sensor-parent',{
    rows:[r({...AHU}),r({id:'B1-RTS-1041',desc:'Room temperature sensor',parent:sys('104'),upn:'104',disc:FMS,im:'VF_I&C_TRANSMITTER'})],
    focus:[{row:1,field:'closestParent'}],
    caption:'A room sensor parented to the area system rather than the equipment it controls.',fix:'Parent B1-RTS-1041 to B1-AHU-1041.'}),
]));

export const AUDIT_EXAMPLE_FIELD_LABELS=Object.freeze(Object.fromEntries(EXTO_REV21_COLUMNS.map(column=>[column.field,column.header])));
const AUDIT_EXAMPLE_COLUMN_ORDER=Object.freeze(['equipmentId','equipmentDescription','closestParent','closestParentStatus','dependencies','dependencyProject','upn','systemName','discipline','building','itemMaster','milestoneParent','milestone','equipmentClassification']);

/* The columns worth showing for one example: every column any row fills in,
   plus every focused column, in registry order. */
export function auditExampleColumns(example){
  const focused=new Set(example.focus.map(cell=>cell.field)),wanted=new Set(focused);
  for(const row of example.rows)for(const field of AUDIT_EXAMPLE_COLUMN_ORDER)if(String(row[field]||'').trim())wanted.add(field);
  /* Closest Parent Status is implied (New) on every system root; it only earns a
     column when a row's status says something the parent does not already. */
  if(!focused.has('closestParentStatus')&&!example.rows.some(row=>String(row.closestParentStatus||'').trim()&&row.closestParent!==row.systemName))wanted.delete('closestParentStatus');
  const lead=['equipmentId','equipmentDescription'].filter(field=>wanted.has(field));
  const flagged=AUDIT_EXAMPLE_COLUMN_ORDER.filter(field=>focused.has(field)&&!lead.includes(field));
  const rest=AUDIT_EXAMPLE_COLUMN_ORDER.filter(field=>wanted.has(field)&&!lead.includes(field)&&!focused.has(field));
  return [...lead,...flagged,...rest];
}

/* A snapshot the engine can run: every VF Exto Upload Template field present, plus the _source
   the engine expects on a row. */
export function auditExampleSnapshot(example){
  const rows=example.rows.map((row,index)=>{
    const record={};
    for(const column of EXTO_REV21_COLUMNS)record[column.field]=String(row[column.field]||'');
    record._source=Object.freeze({file:'example',sheet:'Example',row:index+2,columns:Object.freeze({})});
    return Object.freeze(record);
  });
  return Object.freeze({rows:Object.freeze(rows)});
}
