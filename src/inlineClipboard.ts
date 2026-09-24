import {getBodySlice,type IDocumentData} from '@univerjs/core'
import {validateInlineDocument} from './inlineMedia'

export const INLINE_CLIPBOARD_LIMITS=Object.freeze({bytes:2*1024*1024,cells:10_000,axis:10_000})
export type InlineClipboardFragment={version:1;rows:number;columns:number;documents:[number,number,IDocumentData][]}
/** Clipboard is untrusted, even when another instance of this package wrote it. */
export function decodeInlineClipboard(encoded:string):InlineClipboardFragment{
  if(encoded.length>INLINE_CLIPBOARD_LIMITS.bytes*3)throw new Error('INLINE_CLIPBOARD_TOO_LARGE')
  const json=decodeURIComponent(encoded)
  if(new TextEncoder().encode(json).length>INLINE_CLIPBOARD_LIMITS.bytes)throw new Error('INLINE_CLIPBOARD_TOO_LARGE')
  const value=JSON.parse(json,(key,v)=>{if(['__proto__','constructor','prototype'].includes(key))throw new Error('UNSAFE_CLIPBOARD_KEY');return v}) as InlineClipboardFragment
  if(value?.version!==1||!Number.isInteger(value.rows)||!Number.isInteger(value.columns)||value.rows<1||value.columns<1||value.rows>10_000||value.columns>10_000||value.rows*value.columns>10_000||!Array.isArray(value.documents)||value.documents.length>10_000)throw new Error('INVALID_INLINE_CLIPBOARD')
  const seen=new Set<string>()
  for(const entry of value.documents){
    if(!Array.isArray(entry)||entry.length!==3)throw new Error('INVALID_INLINE_CLIPBOARD')
    const [r,c,doc]=entry,key=`${r}:${c}`
    if(!Number.isInteger(r)||!Number.isInteger(c)||r<0||c<0||r>=value.rows||c>=value.columns||seen.has(key)||!doc?.body||typeof doc.body.dataStream!=='string')throw new Error('INVALID_INLINE_CLIPBOARD')
    seen.add(key);validateInlineDocument(doc)
  }
  return value
}
export function encodeInlineClipboard(fragment:InlineClipboardFragment){
  const encoded=encodeURIComponent(JSON.stringify(fragment));decodeInlineClipboard(encoded);return encoded
}
/** A copied single cell can be pasted into a text cursor, not just another
 * whole cell. Strip only the document terminator; keep paragraphs and objects. */
export function inlineTextFragment(fragment:InlineClipboardFragment):IDocumentData|null{
  if(fragment.rows!==1||fragment.columns!==1||fragment.documents.length!==1)return null
  const doc=structuredClone(fragment.documents[0][2]),body=doc.body!
  if(body.dataStream.endsWith('\r\n'))doc.body=getBodySlice(body,0,body.dataStream.length-2)
  validateInlineDocument(doc);return doc
}
