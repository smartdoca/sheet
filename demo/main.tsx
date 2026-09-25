import React, { useState } from 'react'
import ReactDOM from 'react-dom/client'

import {
  SpreadsheetEditor,
  type WorkbookPersistenceAdapter,
  type WorkbookSnapshot,
} from '../src'
import './style.css'

const STORAGE_KEY = 'univer-sheet-mvp-demo'

const localPersistence: WorkbookPersistenceAdapter = {
  async load() {
    const value = localStorage.getItem(STORAGE_KEY)
    return value ? { snapshot: JSON.parse(value) as WorkbookSnapshot } : null
  },
  async save(snapshot) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot))
  },
}

function App() {
  const [locale, setLocale] = useState('zh')
  return (
    <main className="uos-demo">
      <nav className="uos-demo__nav">
        <a href="?acceptance=index">打开交互验收目录</a>
        <span>隔离测试文档，不修改当前工作簿</span>
        <button type="button" aria-pressed={locale === 'zh'} onClick={() => setLocale('zh')}>中文</button>
        <button type="button" aria-pressed={locale === 'en'} onClick={() => setLocale('en')}>English</button>
      </nav>
      <section className="uos-demo__frame">
        <SpreadsheetEditor
          locale={locale}
          workbookId="demo-workbook"
          workbookName="示例工作簿"
          persistence={localPersistence}
          onError={(error) => console.error(error)}
        />
      </section>
    </main>
  )
}

const acceptanceRoom = new URLSearchParams(location.search).get('host-test')
if (new URLSearchParams(location.search).has('acceptance')) {
  void import('../examples/acceptance-suite').then(m=>m.mountAcceptanceSuite(document.getElementById('root')!))
} else if (new URLSearchParams(location.search).has('inline-test')) {
  void import('../examples/inline-acceptance').then(m=>m.createInlineAcceptance(new URLSearchParams(location.search).get('inline-test')!)).then(App=>ReactDOM.createRoot(document.getElementById('root')!).render(<App/>))
} else if (new URLSearchParams(location.search).has('toolbar-test')) {
  void import('./toolbar-test').then(({ mountToolbarTest }) => mountToolbarTest())
} else if (new URLSearchParams(location.search).has('comment-test')) {
  void import('./comment-test').then(({ mountCommentTest }) => mountCommentTest(new URLSearchParams(location.search).get('comment-test')!))
} else if (new URLSearchParams(location.search).has('border-test')) {
  void import('./border-test').then(({ mountBorderTest }) => mountBorderTest())
} else if (acceptanceRoom) {
  void import('./host-test').then(({ mountHostTest }) => mountHostTest(acceptanceRoom))
} else ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
