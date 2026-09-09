import { clean } from '../core/text.js'

export function auditSparrowMilestoneMigration(){
  const groups=[
    [['122','123'],'123','Utility and Sub Utility Trench Levels are HPM Ready'],
    [['117','115'],'124','50% North side (AB-AR) ready for TI Prefac, PCD (UL & SF) and AMHS start'],
    [['130_1','130'],'130_1','100% Ready for Prefac Start and AMHS Install'],
    [['125'],'125','PSSS Rooms, BGY are HPM Ready'],
    [['128'],'128','50% CR is ready for Data Collection start'],
    [['127','126'],'126','50% North side (AB-AR) ready for First Tool Move-In, Set and sign off - partial HPM ready'],
    [['132'],'128_1','100% CR is ready for Data Collection start'],
    [['133','131'],'133','100% FAB, SubFab, Link and TTN are HPM Ready and ready for Tools Set'],
  ];
  return auditReadMilestoneMigration({format:'ssm-audit-milestone-map',version:1,project:'Sparrow',mappings:groups.flatMap(([from,to,title])=>from.map(code=>({
    from:`SP-L1-M1-${code}`,to:`SP-L1-M1-${to}`,label:`SP-L1-M1-${to} - Sparrow Mod 1 - Level 1 Milestone - ${title}`,aliases:[],
  })))});
}

function auditMigrationCode(value){return clean(value).match(/^((?:[A-Z0-9.]+[-_])*L1(?:[-_][A-Z0-9.]+)+)(?=\s|$)/i)?.[1].toUpperCase()||'';}
export function auditReadMilestoneMigration(value){
  if(value?.format!=='ssm-audit-milestone-map'||value.version!==1||typeof value.project!=='string'||!clean(value.project)||value.project.length>100||!Array.isArray(value.mappings)||!value.mappings.length||value.mappings.length>1000)throw new Error('Select a valid project milestone mapping JSON.');
  const seen=new Map(),aliases=new Map(),targets=new Map(),mappings=[];
  for(const item of value.mappings){
    if(!item||typeof item.from!=='string'||typeof item.to!=='string'||item.from.length>200||item.to.length>200||typeof item.label!=='string'||item.label.length>4000)throw new Error('A milestone mapping has invalid values.');
    const from=auditMigrationCode(item.from),to=auditMigrationCode(item.to),label=clean(item.label);
    if(!from||from!==clean(item.from).toUpperCase()||!to||to!==clean(item.to).toUpperCase()||auditMigrationCode(label)!==to)throw new Error('Each mapping needs exact L1 codes and a replacement label starting with its new code.');
    if(seen.has(from))throw new Error('A milestone code is mapped more than once.');
    if(targets.has(to)&&targets.get(to)!==label)throw new Error('A replacement milestone has conflicting labels.');
    targets.set(to,label);
    const names=item.aliases||[];if(!Array.isArray(names)||names.length>100||names.some(name=>typeof name!=='string'||!clean(name)||name.length>4000))throw new Error('A milestone alias is invalid.');
    const entry={from,to,label,aliases:names.map(clean)};seen.set(from,entry);mappings.push(entry);
    for(const alias of entry.aliases){const key=alias.toUpperCase();if(aliases.has(key)&&aliases.get(key)!==to)throw new Error('A milestone alias has conflicting replacements.');aliases.set(key,to);}
  }
  for(const item of mappings){const next=seen.get(item.to);if(next&&(next.to!==item.to||next.label!==item.label))throw new Error('Chained milestone replacements are not supported. Map directly to the final milestone.');}
  return {format:value.format,version:1,project:clean(value.project),mappings};
}
export function auditReadMigrationSettings(value){
  if(value==null)return {enabled:false,profile:null};
  if(typeof value.enabled!=='boolean')throw new Error('The saved milestone setting is invalid.');
  const profile=value.profile==null?null:auditReadMilestoneMigration(value.profile);
  if(value.enabled&&!profile)throw new Error('A milestone mapping is required before enabling it.');
  return {enabled:value.enabled,profile};
}
export function auditMigrationValue(value,profile){
  const text=clean(value);if(!text||!profile)return null;
  const code=auditMigrationCode(text),entry=profile.mappings.find(item=>code?item.from===code:item.aliases.some(alias=>alias.toUpperCase()===text.toUpperCase()));
  return entry&&entry.label!==text?entry:null;
}
export function auditMigrationReferences(references,settings){
  if(!settings?.enabled||!settings.profile||!references?.milestones)return references;
  const entries=references.milestones.entries.map(entry=>{
    const mapped=auditMigrationValue(entry.parentId,settings.profile)||auditMigrationValue(entry.parentTitle,settings.profile);
    return mapped?{...entry,parentId:mapped.to,parentTitle:mapped.label.slice(mapped.to.length).trim().replace(/^[-:]\s*/,''),milestoneParent:mapped.label}:entry;
  });
  return {...references,milestones:{...references.milestones,entries}};
}
export function auditMilestoneMigrationRows(snapshot,settings){
  if(!settings?.enabled||!settings.profile)return [];
  return snapshot.rows.flatMap(row=>{const mapping=auditMigrationValue(row.milestoneParent,settings.profile);return mapping&&row._source?.columns?.milestoneParent!=null?[{row,mapping}]:[];});
}
export function auditMigrationImpact(impact,before,after,settings){
  if(!settings?.enabled||!settings.profile)return impact;
  const key=row=>JSON.stringify([row._source?.sheet,row._source?.row]);
  const originals=new Map(before.rows.map(row=>[key(row),row])),updates=new Map(after.rows.map(row=>[key(row),row]));
  const conflicts=[],retained=[];
  const unsafe=impact.unsafe.filter(finding=>{
    if(!['reference.milestone-parent-mismatch','milestone.incomplete-pair'].includes(finding.rule.id))return true;
    const id=JSON.stringify([finding.sheet,finding.row]),old=originals.get(id),next=updates.get(id);
    const mapped=old&&auditMigrationValue(old.milestoneParent,settings.profile);
    if(!mapped||!next||next.milestoneParent!==mapped.label||old.milestone!==next.milestone)return true;
    // Renaming an existing L1 does not introduce the unchanged missing L2.
    if(finding.rule.id==='milestone.incomplete-pair'){retained.push(finding);return false;}
    conflicts.push(finding);return false;
  });
  // Explicit project renames do not authorize changing an L2 to hide a conflict.
  return {...impact,unsafe,introduced:impact.introduced.filter(finding=>!retained.includes(finding)),changed:[...(impact.changed||[]),...retained],migrationConflicts:conflicts};
}
