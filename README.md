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
| `onImagePreview` | `SpreadsheetImagePreviewHandler` | Native inline image preview. The host checks access and opens its image viewer. |
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

## Image preview

In edit mode, the first image click selects it and the next click previews it, including inside the native cell draft. Readonly mode previews on the first click. `onImagePreview({drawingId, resource, cell})` receives the existing native drawing ID, `{id, kind: 'image', name}` and the zero-based workbook/sheet/cell location. The host resolves the stable resource ID and opens its viewer; no resolved URL is stored. Without a handler, the package uses the native image gallery. Replacing the callback or toggling readonly does not rebuild the editor. Callback errors reach `onError`; pending errors after disposal are ignored. Selection and preview do not commit the cell draft.

`handle.getImageGallery()` returns detached `{drawingId, resource, placement, sheetId, cell}` entries for UUID-backed inline and floating images across the current workbook, including the active uncommitted draft. Entries follow sheet order, row, column, inline text order, then floating order; unanchored native floating images follow anchored images in their sheet (`cell: null`). Match the clicked occurrence using `drawingId` and its sheet/cell location: repeated uses of an asset remain separate entries. The getter neither writes content nor resolves URLs. Embedded BASE64 images have no host asset ID and are excluded. A custom runtime must implement this capability or the getter throws `IMAGE_GALLERY_UNSUPPORTED`.

## Collaboration

Use `collaboration` for content. Use `remoteSelections` only for cursors.

- Local content transactions enter the host outbox. Remote application, selection, scroll, and resize do not.
- `readOnly` does not publish edits or editing selections.
- `currentSessionId` hides this tab's own cursor and keeps other sessions.
- Comment markers use stable row and column identities. Do not store comments as A1 text.
- Filters are shared: ranges and criteria synchronize, undo and recover with the workbook. Filtered row results are derived locally, including after local cell edits; recalculation does not create a content transaction.

Other entry points:

| Import | Use |
|---|---|
| `@smartdoca/sheet/yjs` | Yjs session, restore, and local transactions. |
| `@smartdoca/sheet/model` | Workbook projection and recovery. |
| `@smartdoca/sheet/xlsx` | `xlsxToSnapshot` and `snapshotToXlsx`. |
