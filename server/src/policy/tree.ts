import { XMLParser } from 'fast-xml-parser'
import { parseXml } from '../../../src/core/xml'
export interface Node { name:string; attrs:Record<string,string>; children:Node[]; text?:string }
export interface PolicyDiagnostic {code:'hidden-stripped'|'hidden-unrecognised';part:string;element:string;attr?:string;value?:string}
export interface Context {parts:Map<string,Node>;diagnostics:PolicyDiagnostic[];handled:WeakSet<Node>;dropPart(path:string):void}
export const local=(name:string)=>name.slice(name.lastIndexOf(':')+1)
export function all(root:Node,name:string):Node[]{const result:Node[]=[];const visit=(n:Node)=>{if(local(n.name)===name)result.push(n);n.children.forEach(visit)};visit(root);return result}
export function attr(n:Node,name:string):string|undefined{return Object.entries(n.attrs).find(([k])=>local(k)===name)?.[1]}
export function remove(root:Node,target:Node):void{root.children=root.children.filter(n=>n!==target);for(const n of root.children)remove(n,target)}
const parser=new XMLParser({preserveOrder:true,ignoreAttributes:false,attributeNamePrefix:'',parseTagValue:false,parseAttributeValue:false,trimValues:false,processEntities:true,htmlEntities:false,cdataPropName:'#cdata',commentPropName:'#comment'})
export function readTree(xml:string):Node{
 // The library validator supplies the same well-formedness/scalar checks as
 // parsing; retain qualified names and ordered text separately for serialization.
 if(/<!DOCTYPE|<!ENTITY/i.test(xml))throw new Error('unsupported-format: XML declarations')
 try { parseXml(xml) } catch (error) {
  // Empty self-closing roots are well-formed but the library's normalized
  // object-only helper rejects their scalar representation. Ordered tree below
  // still requires exactly one real element. All validation errors propagate.
  if (!(error instanceof Error) || error.message !== 'Unexpected XML root shape') throw error
 }
 const convert=(items:any[]):Node[]=>items.flatMap(item=>Object.entries(item).filter(([name])=>name!==':@').map(([name,value]):Node=>{
  if(name==='#text')return{name,attrs:{},children:[],text:String(value)}
  if(name==='#cdata'||name==='#comment')return{name,attrs:{},children:[],text:(value as any[]).map(v=>v['#text']??'').join('')}
  return{name,attrs:{...(item[':@']??{})},children:Array.isArray(value)?convert(value):[]}
 }))
 const root={name:'#document',attrs:{},children:convert(parser.parse(xml))}
 if(root.children.filter(n=>!n.name.startsWith('#')&&!n.name.startsWith('?')).length!==1)throw new Error('unsupported-format: XML root')
 return root
}
const escape=(s:string)=>s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/\r/g,'&#13;')
export function writeTree(n:Node):string{
 if(n.name==='#document')return n.children.map(writeTree).join('')
 if(n.name==='#text')return escape(n.text??'')
 if(n.name==='#cdata')return '<![CDATA['+(n.text??'').replace(/]]>/g,']]]]><![CDATA[>')+']]>'
 if(n.name==='#comment')return '<!--'+(n.text??'')+'-->'
 const attrs=Object.entries(n.attrs).map(([k,v])=>` ${k}="${escape(String(v)).replace(/\n/g,'&#10;').replace(/\t/g,'&#9;')}"`).join('')
 if(n.name.startsWith('?'))return `<${n.name}${attrs}?>`
 return n.children.length?`<${n.name}${attrs}>${n.children.map(writeTree).join('')}</${n.name}>`:`<${n.name}${attrs}/>`
}
