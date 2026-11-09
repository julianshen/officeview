/// <reference path="../../src/types/pako-zlib.d.ts" />
import JSZip from 'jszip'
import {inflateRawSync} from 'node:zlib'
import {posix} from 'node:path'
import {OfficePackage} from '../../src/core/zip'
import {all,attr,readTree,writeTree,type Context,type Node,type PolicyDiagnostic} from './policy/tree'
import {projectDocx} from './policy/docx'
import {projectPptx} from './policy/pptx'
import {projectXlsx,purgeUnusedSharedStrings} from './policy/xlsx'
import {scanVisibility} from './policy/visibility-scan'
export type {PolicyDiagnostic} from './policy/tree'
export type Format='docx'|'xlsx'|'pptx'
export class PolicyError extends Error {readonly code='unsupported-format';constructor(cause?:unknown){super(`unsupported-format${cause instanceof Error?`: ${cause.message}`:''}`)}}

/** Synchronous, bounded inspection of the original ZIP directory. Admission
 * reads only content types; projection subsequently validates all XML parts. */
export function admit(bytes:Uint8Array):{ok:true;format:Format}|{ok:false;code:'unsupported-format'}{
 try{
  const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),u16=(p:number)=>view.getUint16(p,true),u32=(p:number)=>view.getUint32(p,true)
  let end=-1;for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(u32(i)===0x06054b50&&i+22+u16(i+20)===bytes.length){end=i;break}
  if(end<0||u16(end+4)!==0||u16(end+6)!==0||u16(end+8)!==u16(end+10))throw Error('zip')
  const count=u16(end+10),size=u32(end+12);let at=u32(end+16);const limit=at+size
  if(count===65535||limit>end)throw Error('zip64/directory')
  const entries=new Map<string,{offset:number;size:number;method:number;flags:number}>(),decoder=new TextDecoder('utf-8',{fatal:true})
  for(let i=0;i<count;i++){
   if(at+46>limit||u32(at)!==0x02014b50)throw Error('directory')
   const len=u16(at+28),extra=u16(at+30),comment=u16(at+32),next=at+46+len+extra+comment
   if(next>limit)throw Error('directory bounds')
   const name=decoder.decode(bytes.subarray(at+46,at+46+len))
   if(entries.has(name)||name.includes('\\')||name.startsWith('/')||name.split('/').includes('..'))throw Error('ambiguous path')
   entries.set(name,{offset:u32(at+42),size:u32(at+20),method:u16(at+10),flags:u16(at+8)});at=next
  }
  if(at!==limit||entries.has('mimetype'))throw Error('not OOXML')
  const entry=entries.get('[Content_Types].xml');if(!entry||entry.flags&1)throw Error('no content types')
  const p=entry.offset;if(u32(p)!==0x04034b50)throw Error('local header')
  const start=p+30+u16(p+26)+u16(p+28);if(start+entry.size>limit)throw Error('part bounds')
  const compressed=bytes.subarray(start,start+entry.size),cap=1024*1024
  const data=entry.method===8?inflateRawSync(compressed,{maxOutputLength:cap}):entry.method===0&&compressed.length<=cap?compressed:undefined
  if(!data)throw Error('compression')
  const ct=readTree(decoder.decode(data)),formats:Array<[Format,string,string]>=[['docx','word/document.xml','wordprocessingml.document.main+xml'],['xlsx','xl/workbook.xml','spreadsheetml.sheet.main+xml'],['pptx','ppt/presentation.xml','presentationml.presentation.main+xml']]
  const matches=formats.filter(([,path,type])=>entries.has(path)&&all(ct,'Override').some(n=>attr(n,'PartName')===`/${path}`&&attr(n,'ContentType')===`application/vnd.openxmlformats-officedocument.${type}`))
  if(matches.length!==1)throw Error('main part');return{ok:true,format:matches[0][0]}
 }catch{return{ok:false,code:'unsupported-format'}}
}

export async function projectPackage(pkg:OfficePackage,format:Format,contentPolicy:'visible-only'|'full'):Promise<{pkg:OfficePackage;contentPolicyId:string;diagnostics:PolicyDiagnostic[]}>{
 if(contentPolicy==='full')return{pkg,contentPolicyId:'full@1',diagnostics:[]}
 if(contentPolicy!=='visible-only'||!['docx','xlsx','pptx'].includes(format))throw new PolicyError()
 try{
  const parts=new Map<string,Node>(),original=new Map<string,Uint8Array>(),initial=new Map<string,string>(),removed=new Set<string>(),diagnostics:PolicyDiagnostic[]=[]
  let remaining=96*1024*1024
  for(const path of pkg.paths()){
   const bytes=await pkg.bytes(path,remaining);if(!bytes)throw Error('missing part');remaining-=bytes.byteLength;original.set(path,bytes)
   if(/\.(?:xml|rels)$/i.test(path)){
    const tree=readTree(new TextDecoder('utf-8',{fatal:true}).decode(bytes));parts.set(path,tree);initial.set(path,writeTree(tree))
   }
  }
  const ctx:Context={parts,diagnostics,handled:new WeakSet(),dropPart(path){removed.add(path);parts.delete(path);const rel=posix.join(posix.dirname(path),'_rels',posix.basename(path)+'.rels');removed.add(rel);parts.delete(rel)}}
  // customXml/ holds data-store items (e.g. SharePoint schema metadata with
  // `ma:hidden`); no painter reads them, so they never travel (§6.4).
  for(const path of [...original.keys()])if(/^customXml\//i.test(path))ctx.dropPart(path)
  if(format==='docx')projectDocx(ctx);else if(format==='xlsx')projectXlsx(ctx);else projectPptx(ctx)
  // Delete relationships/content-types to removed parts, not merely their bytes.
  for(const [path,tree]of parts){
   if(path.endsWith('.rels')){
    const base=path==='_rels/.rels'?'':posix.dirname(path).replace(/(?:^|\/)\_rels$/,'')
    for(const n of all(tree,'Relationship')){if(attr(n,'TargetMode')==='External')continue;const target=attr(n,'Target')??'',resolved=target.startsWith('/')?target.slice(1):posix.normalize(posix.join(base,target))
     if(removed.has(resolved)){const prune=(root:Node)=>{root.children=root.children.filter(c=>c!==n);root.children.forEach(prune)};prune(tree)}
    }
   }
   if(path==='[Content_Types].xml')tree.children.forEach(root=>{root.children=root.children.filter(n=>!removed.has((attr(n,'PartName')??'').replace(/^\//,'')))})
  }
  scanVisibility(ctx)
  if(format==='xlsx')purgeUnusedSharedStrings(ctx)
  const zip=new JSZip()
  for(const [path,bytes]of original){if(removed.has(path))continue;const tree=parts.get(path),serialized=tree&&writeTree(tree)
   // Untouched XML remains byte-identical, including declarations and comments.
   zip.file(path,serialized!==undefined&&serialized!==initial.get(path)?serialized:bytes)
  }
  const out=await OfficePackage.load(await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}))
  return{pkg:out,contentPolicyId:'visible-only@1',diagnostics}
 }catch(error){if(error instanceof PolicyError)throw error;throw new PolicyError(error)}
}
