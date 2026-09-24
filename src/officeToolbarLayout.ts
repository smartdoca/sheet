/** Widths include each group's padding and divider; keep aligned with compact CSS. */
export const OFFICE_GROUP_WIDTHS = {edit:232,insert:54,font:216,align:88,layout:80,number:144,data:256,formula:54} as const
export function officeToolbarLayout<T extends {width:number}>(groups:T[],available:number,hostWidth:number){
  const overflow=groups.reduce((sum,g)=>sum+g.width,0)+hostWidth>available
  const align=overflow||available<1400?'start':'center'
  if(!overflow)return {visible:groups,hidden:[] as T[],align}
  const budget=Math.max(0,available-hostWidth-56)
  let used=0,count=0
  for(const group of groups){if(used+group.width>budget)break;used+=group.width;count++}
  return {visible:groups.slice(0,count),hidden:groups.slice(count),align}
}
