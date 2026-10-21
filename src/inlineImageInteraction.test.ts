import {expect,it,vi} from 'vitest'
import {attachInlineImageInteraction,type InlineImageHit} from './inlineImageInteraction'

const image:InlineImageHit={key:'book:s:2:3:drawing:asset',drawingId:'drawing',resource:{id:'asset',kind:'image',name:'picture.png'},cell:{workbookId:'book',sheetId:'s',row:2,column:3}}
const event=(type:string,extra:Record<string,unknown>={})=>Object.assign(new Event(type,{cancelable:true}),{button:0,clientX:20,clientY:30,...extra})
function fixture(){
  const target=new EventTarget(),select=vi.fn(),preview=vi.fn(),onError=vi.fn(),refresh=vi.fn()
  let hit:InlineImageHit|null=structuredClone(image),readonly=false
  const adapter=attachInlineImageInteraction(target,{hit:()=>hit,readOnly:()=>readonly,select,preview,refresh,onError})
  return {adapter,select,preview,onError,target,setHit:(next:InlineImageHit|null)=>{hit=next},readonly:()=>{readonly=true},
    click(extra:Record<string,unknown>={}){target.dispatchEvent(event('pointerdown',extra));target.dispatchEvent(event('click',extra))},
  }
}
it('selects first in edit mode, previews the selected occurrence next, and resets on another target',async()=>{
  const f=fixture()
  try{
    f.click();await Promise.resolve()
    expect(f.preview).not.toHaveBeenCalled();expect(f.adapter.isSelected('drawing')).toBe(true)
    f.click();await Promise.resolve()
    expect(f.preview).toHaveBeenCalledExactlyOnceWith({drawingId:image.drawingId,resource:image.resource,cell:image.cell})
    f.setHit(null);f.click();expect(f.adapter.isSelected('drawing')).toBe(false)
    f.setHit({...image,key:'second',drawingId:'second'});f.click();await Promise.resolve()
    expect(f.preview).toHaveBeenCalledTimes(1)
  }finally{f.adapter.dispose()}
})
it('selects the image glyph in an active draft without writing content, and previews readonly on the first click',async()=>{
  const f=fixture()
  try{
    const draft={...image,draft:{unitId:'native-draft',startOffset:4,endOffset:5}}
    f.setHit(draft);f.click();await Promise.resolve()
    expect(f.select).toHaveBeenCalledExactlyOnceWith(draft);expect(f.preview).not.toHaveBeenCalled()
    f.click();await Promise.resolve();expect(f.preview).toHaveBeenCalledTimes(1)
    f.adapter.clear();f.readonly()
    await Promise.resolve();await Promise.resolve();f.click();await Promise.resolve()
    expect(f.preview).toHaveBeenCalledTimes(2)
  }finally{f.adapter.dispose()}
})
it('does not preview modifier clicks, drags, changed assets or disposed callbacks; reports host errors',async()=>{
  const f=fixture()
  try{
    for(const modifier of ['shiftKey','ctrlKey','metaKey','altKey'])f.click({[modifier]:true})
    f.target.dispatchEvent(event('pointerdown'));f.target.dispatchEvent(event('click',{clientX:40}))
    f.target.dispatchEvent(event('pointerdown'));f.setHit({...image,key:'changed-asset'});f.target.dispatchEvent(event('click'))
    await Promise.resolve();expect(f.preview).not.toHaveBeenCalled()
    f.readonly();f.preview.mockImplementation(()=>{throw new Error('Access denied')})
    f.click();await Promise.resolve();await Promise.resolve()
    expect(f.onError).toHaveBeenCalledWith(new Error('Access denied'))
    await Promise.resolve();f.click();f.adapter.dispose();await Promise.resolve()
    expect(f.preview).toHaveBeenCalledTimes(1)
  }finally{f.adapter.dispose()}
})
