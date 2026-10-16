import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

/* The build concatenates modules into one classic-script scope, where a later
   top-level declaration silently replaces an earlier one with the same name.
   Returns every name declared at the top level of more than one module. */
const TOP_LEVEL=/^(?:export\s+)?(?:async\s+)?(?:function\s*\*?\s*([A-Za-z_$][\w$]*)|(?:const|let|var|class)\s+([A-Za-z_$][\w$]*))/gm
export function duplicateTopLevelNames(root,paths){
  const owner=new Map(),duplicates=[]
  for(const path of paths){
    for(const match of readFileSync(resolve(root,path),'utf8').matchAll(TOP_LEVEL)){
      const name=match[1]||match[2]
      if(owner.has(name)&&owner.get(name)!==path)duplicates.push(`${name} (${owner.get(name)} and ${path})`)
      else owner.set(name,path)
    }
  }
  return duplicates
}
