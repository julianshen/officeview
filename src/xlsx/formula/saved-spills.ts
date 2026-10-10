import type { XlsxCell, XlsxDocument, XlsxSheet } from '../types'

export interface SavedSpillFollower { sheet:XlsxSheet;cell:XlsxCell;anchor:XlsxCell;written:Omit<XlsxCell,'ref'|'col'|'row'> }
const followers=new WeakMap<XlsxDocument,SavedSpillFollower[]>()
function lowerBound<T>(items:readonly T[],value:number,key:(item:T)=>number):number{
  let lo=0,hi=items.length
  while(lo<hi){const mid=(lo+hi)>>>1;if(key(items[mid])<value)lo=mid+1;else hi=mid}
  return lo
}
function address(ref:string):{row:number;col:number}|undefined{
  const m=/^([A-Z]+)([1-9]\d*)$/.exec(ref);if(!m)return undefined
  let col=0;for(const char of m[1])col=col*26+char.charCodeAt(0)-64
  const row=Number(m[2]);return col<=16384&&row<=1048576?{row:row-1,col:col-1}:undefined
}
/** Capture actual parser-owned objects once, before callers can edit/replace
 * them. Merely passing a lookalike cell cannot acquire a lease over user data. */
export function initializeSavedSpills(doc:XlsxDocument):void{
  if(followers.has(doc))return
  const result:SavedSpillFollower[]=[],seen=new Map<XlsxCell,SavedSpillFollower|null>()
  for(const sheet of doc.sheets){
    const rows=[...sheet.rows].sort((a,b)=>a.index-b.index)
    const sortedCells=new Map(rows.map(row=>[row,[...row.cells].sort((a,b)=>a.col-b.col)]))
    for(const row of sheet.rows)for(const anchor of row.cells){
      if(!anchor.dynamicArray||!anchor.arrayRef)continue
      const refs=anchor.arrayRef.split(':');if(refs.length>2)continue
      const first=address(refs[0]),last=address(refs[1]??refs[0]);if(!first||!last||first.col!==anchor.col||first.row!==anchor.row||last.col<first.col||last.row<first.row)continue
      if(first.row===last.row&&first.col===last.col)continue
      for(let ri=lowerBound(rows,first.row,row=>row.index);ri<rows.length&&rows[ri].index<=last.row;ri++){
        const row=rows[ri],cells=sortedCells.get(row)!
        for(let ci=lowerBound(cells,first.col,cell=>cell.col);ci<cells.length&&cells[ci].col<=last.col;ci++){
          const cell=cells[ci]
          if(cell===anchor||cell.col<first.col||cell.col>last.col||cell.formula!==undefined||cell.sharedFormula!==undefined||cell.hasCachedValue!==true)continue
          const {ref:_ref,col:_col,row:_row,...written}=cell
          if(seen.has(cell)){seen.set(cell,null);continue}
          seen.set(cell,{sheet,cell,anchor,written})
        }
      }
    }
  }
  for(const record of seen.values())if(record)result.push(record)
  followers.set(doc,result)
}
export function savedSpillFollowers(doc:XlsxDocument):readonly SavedSpillFollower[]{return followers.get(doc)??[]}
