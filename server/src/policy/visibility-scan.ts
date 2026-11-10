import {attr,local,type Context,type Node} from './tree'
const drawable=new Set(['r','wsp','wgp','sp','pic','grpSp','graphicFrame','cxnSp','drawing','c','row','oneCellAnchor','twoCellAnchor','absoluteAnchor'])
function unsafe(key:string,value:string):boolean{
 const v=value.trim().toLowerCase()
 if(['hidden','vanish','specVanish','webHidden'].includes(key))return !['0','false','off'].includes(v)
 if(['show','visible'].includes(key))return !['1','true','on','show','visible'].includes(v)
 if(['visibility','state'].includes(key))return !['visible','show'].includes(v)
 return false
}
/** Unknown flags never survive by merely deleting the flag: remove the owning
 * drawable, or refuse the part if a safe drawable boundary cannot be found. */
export function scanVisibility(ctx:Context):void{
 for(const [part,root]of ctx.parts){const removals=new Set<Node>()
  const visit=(n:Node,ancestors:Node[])=>{
   if(!ctx.handled.has(n)){
    const flags=Object.entries(n.attrs).filter(([k,v])=>k!=='xmlns'&&!k.startsWith('xmlns:')&&unsafe(local(k),v))
    if(['vanish','specVanish','webHidden'].includes(local(n.name))&&unsafe(local(n.name),attr(n,'val')??'1'))flags.push([n.name,attr(n,'val')??'1'])
    for(const [key,value]of flags){const target=[...ancestors,n].reverse().find(a=>drawable.has(local(a.name)))
     if(!target)throw new Error(`unsupported-format: unowned visibility flag ${n.name} ${key}=${value} in ${part}`)
     removals.add(target);ctx.diagnostics.push({code:'hidden-unrecognised',part,element:n.name,attr:key,value})
    }
   }
   for(const child of n.children)visit(child,[...ancestors,n])
  };visit(root,[])
  const prune=(n:Node)=>{n.children=n.children.filter(c=>!removals.has(c));n.children.forEach(prune)};prune(root)
 }
}
