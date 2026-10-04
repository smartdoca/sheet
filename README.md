# @smartdoca/sheet

[中文](README.zh-CN.md)

Embeddable collaborative spreadsheet editor for React, powered by Univer Sheets. The package owns the workbook model and grid. The host owns identity, files, permissions, and the network.

Licensed under [MIT](LICENSE).

## Install

```sh
npm install @smartdoca/sheet react react-dom
```

```tsx
import { SpreadsheetEditor } from "@smartdoca/sheet";
import "@smartdoca/sheet/style.css";

export function Sheet({ id }: { id: string }) {
  return <SpreadsheetEditor workbookId={id} readOnly={false} />;
}
```

## Props

`SpreadsheetEditor` accepts `SpreadsheetEditorProps`. `workbookId` is required.

| Prop | Type | Role |
|---|---|---|
| `workbookId` | `string` | Stable workbook id. |
| `workbookName` | `string` | Display name. |
| `initialSnapshot` | `WorkbookSnapshot` | Snapshot applied when the editor mounts. |
| `readOnly` | `boolean` | Stops cell editing. |
| `collaboration` | `CollaborationAdapter` | Host bridge for workbook changes. |
| `persistence` | `WorkbookPersistenceAdapter` | Optional local persistence. Do not add a second autosave path beside the host. |
| `resourceAdapter` | `ResourceAdapter` | Host file and image storage. |
| `onAttachmentPreview` | `SpreadsheetAttachmentPreviewHandler` | Inline attachment label clicks, including readonly. The host opens the preview. |
| `remoteSelections` | `SpreadsheetRemoteSelection[]` | Ephemeral selections. Identity and color come from the host session. |
| `currentSessionId` | `string` | This tab's session. Other tabs of the same user stay visible. |
| `commentMarkers` | `SpreadsheetCommentMarker[]` | Host-owned comments on stable anchors. |
| `activeCommentId` | `string \| null` | Highlighted comment. |
| `onCommentAnchorClick` | function | One marker was activated. |
| `onCommentAnchorsClick` | function | Every overlapping marker. There is no implicit first-comment choice. |
| `onSelectionChange` | function | Local cell selection. |
| `onChange` | `(snapshot) => void` | Workbook snapshot after a local change. |
| `onReady` | `(handle) => void` | Receives `SpreadsheetEditorHandle`. |
| `onSaveStateChange` | function | Host save indicator. |
| `locale` | `SpreadsheetLocale` | `zh` and `zh-*` stay Chinese. Any other code shows English. Omitted stays Chinese. |
| `messages` | `Record<string, string>` | Replaces individual message keys. |
| `autoSave` | `boolean` | Built-in save timer. Leave it off when the host owns persistence. |
| `showSaveState` | `boolean` | Hides the package save badge when the host draws its own. |
| `showHeader` | `boolean` | Workbook header. |
| `showInsertToolbar` | `boolean` | Insertion row, independent of the document header. |
| `toolbarLayout` | `SpreadsheetToolbarLayout` | `simple` flattens native categories. |
| `className`, `style`, `classNames`, `styles` | | Root and region styling. |

`initialRows`, `initialColumns`, `autoFitContent`, menus, image upload, and inline actions are also on `SpreadsheetEditorProps`.

Updating props, readonly, or selection must not rebuild the workbook.

## Attachment preview

```tsx
<SpreadsheetEditor
  workbookId={id}
  onAttachmentPreview={async ({ node, cell }) => {
    // Resolve node.refId with the host's current asset permissions.
    await openAttachmentPreview(node.refId, cell);
  }}
  onError={reportError}
/>
```

`SpreadsheetAttachmentPreviewEvent` contains `phase: 'click'`, `node: { type: 'attachment', refId, label }`, and `cell: { workbookId, sheetId, row, column }` (zero-based coordinates). Only clicks on the native inline attachment label trigger it; hover and clicks elsewhere in the cell do not. Preview works in readonly mode and does not change workbook content. Replacing or removing the callback does not rebuild the editor. Synchronous throws and rejected promises reach the latest `onError`; pending errors are ignored after disposal. No callback means no preview action. Doca owns metadata lookup, access checks, preview URLs and UI.

## Collaboration

Use `collaboration` for content. Use `remoteSelections` only for cursors.

- Local content transactions enter the host outbox. Remote application, selection, scroll, and resize do not.
- `readOnly` does not publish edits or editing selections.
- `currentSessionId` hides this tab's own cursor and keeps other sessions.
- Comment markers use stable row and column identities. Do not store comments as A1 text.

Other entry points:

| Import | Use |
|---|---|
| `@smartdoca/sheet/yjs` | Yjs session, restore, and local transactions. |
| `@smartdoca/sheet/model` | Workbook projection and recovery. |
| `@smartdoca/sheet/xlsx` | `xlsxToSnapshot` and `snapshotToXlsx`. |
