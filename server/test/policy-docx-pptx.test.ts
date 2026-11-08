import { describe, expect, test } from 'bun:test'
import { projectDocx } from '../src/policy/docx'
import { projectPptx } from '../src/policy/pptx'
import type { Context, Node } from '../src/policy/tree'
const n = (name:string, attrs:Record<string,string> = {}, children:Node[] = []):Node => ({name,attrs,children})
const run = (value:string, props:Node[] = []) => n('w:r', {label:value}, [n('w:rPr',{},props), n('w:t',{},[{name:'#text',attrs:{},children:[],text:value}])])
function context(parts:Record<string,Node>):Context { const map=new Map(Object.entries(parts)); return {parts:map,diagnostics:[],handled:new WeakSet(),dropPart(path){map.delete(path)}} }
const labels = (node:Node):string[] => [node.attrs.label ?? '', ...node.children.flatMap(labels)].filter(Boolean)
describe('DOCX effective hidden content projection',()=>{
 test('defaults, paragraph and character style chains, direct false overrides',()=>{
  const styles=n('w:styles',{},[
   n('w:docDefaults',{},[n('w:rPrDefault',{},[n('w:rPr',{},[n('w:vanish')])])]),
   n('w:style',{'w:styleId':'base','w:type':'paragraph'},[n('w:rPr',{},[n('w:vanish',{'w:val':'0'})])]),
   n('w:style',{'w:styleId':'para','w:type':'paragraph'},[n('w:basedOn',{'w:val':'base'})]),
   n('w:style',{'w:styleId':'secret','w:type':'character'},[n('w:rPr',{},[n('w:webHidden')])]),
   n('w:style',{'w:styleId':'derived','w:type':'character'},[n('w:basedOn',{'w:val':'secret'})])])
  const doc=n('w:document',{},[n('w:p',{},[run('default-secret')]),n('w:p',{},[n('w:pPr',{},[n('w:pStyle',{'w:val':'para'})]),run('visible'),run('style-secret',[n('w:rStyle',{'w:val':'derived'})]),run('override',[n('w:rStyle',{'w:val':'derived'}),n('w:webHidden',{'w:val':'false'})])])])
  const ctx=context({'word/document.xml':doc,'word/styles.xml':n('#document',{},[styles])});projectDocx(ctx)
  expect(labels(doc)).toEqual(['visible','override']);expect(ctx.diagnostics.length).toBe(2)
 })
 test('a hidden paragraph mark is recognised: no refusal, its runs keep their own cascade',()=>{
  const flag=n('w:vanish'),doc=n('w:document',{},[n('w:p',{},[n('w:pPr',{},[n('w:rPr',{},[flag])]),run('kept')])])
  const ctx=context({'word/document.xml':doc});projectDocx(ctx)
  expect(labels(doc)).toEqual(['kept']);expect(ctx.handled.has(flag)).toBe(true)
  const odd=n('w:vanish',{'w:val':'maybe'}),doc2=n('w:document',{},[n('w:p',{},[n('w:pPr',{},[n('w:rPr',{},[odd])])])])
  const ctx2=context({'word/document.xml':doc2});projectDocx(ctx2);expect(ctx2.handled.has(odd)).toBe(false)
 })
 test('hidden drawing run and group child are purged; comments are dropped',()=>{
  const hidden=n('w:r',{label:'drawing-secret'},[n('w:drawing',{},[n('wp:docPr',{hidden:'true'})])])
  const group=n('wpg:wgp',{},[n('wps:wsp',{label:'group-secret'},[n('wps:cNvPr',{hidden:'1'})]),n('pic:pic',{label:'group-visible'},[n('pic:cNvPr',{hidden:'false'})])])
  const doc=n('w:document',{},[hidden,group,n('w:commentReference',{'w:id':'1'})]);const ctx=context({'word/document.xml':doc,'word/comments.xml':n('w:comments')});projectDocx(ctx)
  expect(labels(doc)).toEqual(['group-visible']);expect(ctx.parts.has('word/comments.xml')).toBe(false);expect(doc.children.some(x=>x.name==='w:commentReference')).toBe(false)
 })
})
describe('PPTX projection',()=>{
 test('drops hidden slides by slide flag or slide ID flag plus notes/comments',()=>{
  const presentation=n('p:presentation',{},[n('p:sldIdLst',{},[n('p:sldId',{id:'256','r:id':'a'}),n('p:sldId',{id:'257','r:id':'b',show:'0'}),n('p:sldId',{'r:id':'c'})])]);
  const rels=n('Relationships',{},['a','b','c'].map(id=>n('Relationship',{Id:id,Target:`slides/${id}.xml`,Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'})))
  const ctx=context({'ppt/presentation.xml':presentation,'ppt/_rels/presentation.xml.rels':rels,'ppt/slides/a.xml':n('p:sld',{show:'0'}),'ppt/slides/b.xml':n('p:sld'),'ppt/slides/c.xml':n('p:sld',{show:'1'}),'ppt/notesSlides/notesSlide1.xml':n('p:notes'),'ppt/comments/comment1.xml':n('p:cmLst')});projectPptx(ctx)
  expect(presentation.children[0].children.map(x=>x.attrs['r:id'])).toEqual(['c']);expect([...ctx.parts.keys()]).toEqual(['ppt/presentation.xml','ppt/_rels/presentation.xml.rels','ppt/slides/c.xml'])
 })
 test('removes hidden group as a whole and hidden individual shapes',()=>{
  const shape=(label:string,hidden:string)=>n('p:sp',{label},[n('p:nvSpPr',{},[n('p:cNvPr',{hidden})])]);const doc=n('p:sld',{},[n('p:grpSp',{label:'secret-group'},[n('p:nvGrpSpPr',{},[n('p:cNvPr',{hidden:'1'})]),shape('inherited-secret','false')]),shape('secret','true'),shape('visible','false'),shape('unknown','maybe')]);const ctx=context({'ppt/slides/s.xml':doc});projectPptx(ctx);expect(labels(doc)).toEqual(['visible','unknown']);expect(ctx.diagnostics.length).toBe(2)
 })
})

test('DOCX/PPTX Stage-0 fixture XML removes every marker and retains control', async () => {
  const { hiddenFixtures } = await import('../../scripts/stream-s0/hidden-fixtures')
  const { OfficePackage } = await import('../../src/core/zip')
  const { readTree, writeTree } = await import('../src/policy/tree')
  for (const fixture of (await hiddenFixtures()).filter(f => f.kind !== 'xlsx')) {
    const pkg = await OfficePackage.load(fixture.withHidden)
    const parts: Record<string,Node> = {}
    for (const path of pkg.paths(p => /\.(xml|rels)$/.test(p))) parts[path] = readTree((await pkg.text(path))!)
    const ctx = context(parts)
    if (fixture.kind === 'docx') projectDocx(ctx); else projectPptx(ctx)
    const serialized = [...ctx.parts.values()].map(writeTree).join('')
    expect(serialized, fixture.id).not.toContain(fixture.marker)
    expect(serialized, fixture.id).toContain(fixture.control)
  }
})

test('Word default hidden does not reinterpret DrawingML text runs after a Word override', () => {
 const doc=n('w:document',{},[n('w:p',{},[n('w:r',{},[n('w:rPr',{},[n('w:vanish',{'w:val':'0'})]),n('w:drawing',{},[n('a:p',{},[n('a:r',{label:'visible-shape-text'})])])])])])
 const styles=n('w:styles',{},[n('w:docDefaults',{},[n('w:rPrDefault',{},[n('w:rPr',{},[n('w:vanish')])])])])
 projectDocx(context({'word/document.xml':doc,'word/styles.xml':styles}))
 expect(labels(doc)).toEqual(['visible-shape-text'])
})
