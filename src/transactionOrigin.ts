/** Pinned Univer adapter: formula-normalization mutations are derived engine
 * state, not authored content. Do not ignore all onlyLocal commands: native
 * editing may also use that option. */
export function isFormulaProjection(options:unknown):boolean {
  return !!options&&typeof options==='object'&&(options as {fromFormula?:boolean}).fromFormula===true
}
