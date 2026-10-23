/** Native hover debounce flushes its last value when the service completes.
 * Clear it before disposal removes units and completes the hover stream.
 * triggerScroll invalidates the hovered cell without authoring content. */
export function clearHoverBeforeDispose(runtime:{dispose():void},hover:{triggerScroll():void}){
  const dispose=runtime.dispose
  let disposed=false
  runtime.dispose=function(){
    if(disposed)return
    disposed=true
    hover.triggerScroll()
    dispose.call(this)
  }
}
