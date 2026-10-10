/** Resolve dynamic-array provenance through the actual OOXML metadata chain.
 * Prefixes are retained here because similarly named foreign elements must not
 * authorize clearing cached output cells. */
import { XMLParser } from 'fast-xml-parser'
import { parseXml } from '../core/xml'

const MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const DYNAMIC='http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray'
const EXT='{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}'
type Raw={ [key:string]:unknown }
type Node={raw:Raw;ns:Map<string,string>}
function attrs(node:Node):Raw{return (node.raw['@attrs']??{}) as Raw}
function children(node:Node,name:string,uri=MAIN):Node[]{
  const out:Node[]=[]
  for(const [key,value] of Object.entries(node.raw)){
    if(key.startsWith('@')||key.startsWith('#'))continue
    const parts=key.split(':');if(parts.at(-1)!==name)continue
    for(const item of Array.isArray(value)?value:[value]){
      if(!item||typeof item!=='object')continue
      const raw=item as Raw;const ns=new Map(node.ns)
      for(const [key,value] of Object.entries((raw['@attrs']??{}) as Raw)){
        if(key==='xmlns')ns.set('',String(value))
        else if(key.startsWith('xmlns:'))ns.set(key.slice(6),String(value))
      }
      if(ns.get(parts.length===2?parts[0]:'')===uri)out.push({raw,ns})
    }
  }
  return out
}
function index(value:unknown):number|undefined{
  if(typeof value!=='string'||!/^\d+$/.test(value))return undefined
  const n=Number(value);return Number.isSafeInteger(n)?n:undefined
}
function counted(node:Node,name:string):Node[]{
  const result=children(node,name);const count=attrs(node).count
  return count===undefined||index(count)===result.length?result:[]
}
export function dynamicMetadataIndices(xml:string):ReadonlySet<number>{
  parseXml(xml) // Validate entities and XML before using the namespace-aware tree.
  const raw=new XMLParser({ignoreAttributes:false,attributesGroupName:'@attrs',attributeNamePrefix:'',parseAttributeValue:false,parseTagValue:false}).parse(xml) as Raw
  const roots=children({raw,ns:new Map()},'metadata');if(roots.length!==1)return new Set()
  const root=roots[0],typeList=children(root,'metadataTypes'),cellList=children(root,'cellMetadata')
  if(typeList.length!==1||cellList.length!==1)return new Set()
  const types=counted(typeList[0],'metadataType'),cells=counted(cellList[0],'bk')
  const future=children(root,'futureMetadata').filter(n=>attrs(n).name==='XLDAPR')
  if(future.length!==1)return new Set()
  const records=counted(future[0],'bk'),dynamic=new Set<number>()
  records.forEach((record,i)=>{
    const lists=children(record,'extLst');if(lists.length!==1)return
    const extensions=children(lists[0],'ext').filter(n=>String(attrs(n).uri).toLowerCase()===EXT)
    if(extensions.length!==1)return
    const props=children(extensions[0],'dynamicArrayProperties',DYNAMIC)
    if(props.length===1&&['1','true'].includes(String(attrs(props[0]).fDynamic)))dynamic.add(i)
  })
  const result=new Set<number>()
  cells.forEach((cell,i)=>{
    for(const record of children(cell,'rc')){
      const t=index(attrs(record).t),v=index(attrs(record).v)
      if(t===undefined||t<1||t>types.length||v===undefined)continue
      const a=attrs(types[t-1])
      if(a.name==='XLDAPR'&&['1','true'].includes(String(a.cellMeta))&&dynamic.has(v))result.add(i+1)
    }
  })
  return result
}
