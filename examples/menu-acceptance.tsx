import React, { useState } from 'react'
import { createRoot } from 'react-dom/client'
import { SpreadsheetEditor, SPREADSHEET_MENU_PATHS, type WorkbookSnapshot } from '@online-office/univer-sheet'
import { createExlsxBaseline, restoreExlsxDocument, createExlsxCollaborationSession } from '@online-office/univer-sheet/yjs'
import '@online-office/univer-sheet/style.css'

/** Browser fixture: import from the INSTALLED package, not src or a workspace alias. */
export async function mountMenuAcceptance(element: HTMLElement) {
  const base = await createExlsxBaseline({ id: 'menu-acceptance', name: '菜单验收', styles: {}, sheetOrder: ['s'], sheets: { s: {
    id: 's', name: 'Sheet1', rowCount: 200, columnCount: 26, cellData: {},
  } } } as unknown as WorkbookSnapshot, 'menu-epoch')
  const session = await createExlsxCollaborationSession({ doc: await restoreExlsxDocument(base), baseline: base.baseline, sessionId: crypto.randomUUID() })
  const variant = new URLSearchParams(location.search).get('path') ?? 'short-end'
  const paths: Record<string, string | string[]> = {
    'short-end': SPREADSHEET_MENU_PATHS.toolbarEnd,
    'full-end': ['ribbon', 'ribbon.others', 'ribbon.others.others'],
    'pipe-end': 'ribbon|ribbon.others|ribbon.others.others',
    'short-history': 'ribbon.start.history',
  }
  let redraw = () => {}; let writes = 0; let ready = 0
  session.onLocalTransaction(() => { writes++; redraw() })
  function Fixture() {
    const [, update] = useState(0); redraw = () => update(n => n + 1)
    const [narrow, setNarrow] = useState(false)
    const [readonly, setReadonly] = useState(false)
    const [enabled, setEnabled] = useState(true)
    const [visible, setVisible] = useState(true)
    const [clicks, setClicks] = useState(0)
    const [error, setError] = useState('')
    return <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', fontFamily: 'system-ui' }}>
      <div><output aria-label="验收">rc.4 {variant} · ready {ready} · writes {writes} · clicks {clicks}</output>
        <button onClick={() => setNarrow(v => !v)}>切换窄屏</button><button onClick={() => setReadonly(v => !v)}>切换只读</button>
        <button onClick={() => setEnabled(v => !v)}>切换评论权限</button><button onClick={() => setVisible(v => !v)}>切换评论可见</button>
        <output aria-label="错误">{error}</output></div>
      <div style={{ flex: 1, minHeight: 0, width: narrow ? 420 : '100%' }}>
        <SpreadsheetEditor workbookId="menu-acceptance" collaboration={session} readOnly={readonly} showHeader={false} showSaveState={false} showInsertToolbar={false}
          onReady={() => { ready++; redraw() }} onError={error => setError(error.message)}
          menus={[{ id: 'acceptance.comment', title: '评论', ariaLabel: '创建区域评论', path: paths[variant] ?? paths['short-end'], order: 1000, iconOnly: true, tone: 'amber', enabled, visible, requiresEditPermission: false,
            icon: <svg viewBox="0 0 24 24"><path d="M4 4h16v12H9l-5 4z" fill="none" stroke="currentColor" strokeWidth="2" /></svg>,
            action: () => setClicks(n => n + 1),
          }]} />
      </div>
    </div>
  }
  createRoot(element).render(<Fixture />)
}
