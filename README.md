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
| `viewPreferences` | `SpreadsheetViewPreferences` | Personal document-wide `{ zoomRatio }`; `1` is 100%. |
| `onViewPreferencesChange` | `(value: SpreadsheetViewPreferences) => void` | Successful native zoom changes; props and sheet switches do not call it. |
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

## Personal zoom and cell dimensions

```tsx
<SpreadsheetEditor
  workbookId={id}
  viewPreferences={viewPreferences} // Host state, e.g. { zoomRatio: 1 }.
  onViewPreferencesChange={setViewPreferences}
/>
```

`zoomRatio` is a finite number from `0.1` through `4`. One personal scale follows the document across worksheet switches, native additions, remote projection and resize. The native slider, +/- controls, menu and wheel report successful changes, including readonly. Updating props applies silently and does not recreate the editor; a host update during a native cell draft waits until editing ends. The host should update its state before attempting storage so a storage failure retains the current choice. Omitting both props retains existing native behavior; supplying only the callback observes the native initial scale. Removing a callback stops notifications; removing a controlled value retains the mounted view's current personal scale.

The package has no account lookup or browser storage. `isSpreadsheetViewPreferences` from `@smartdoca/sheet/preferences` validates exactly `{ zoomRatio }` without importing React, Univer, or Yjs. The host validates its own record version and scope separately. Doca uses default 100% and a version-1 browser record scoped by origin/account/document, written only after explicit native changes. Missing records do not create records. Unknown/malformed records remain unused and are not overwritten; storage failures retain only the current view choice. Rollback stops consuming records and preserves them.

Personal zoom is excluded from content transactions, undo, autosave and package snapshot/save/export boundaries. Original `zoomRatio` fields, including omission, are preserved there; `getRuntime()` still exposes the native runtime's current display scale. Canonical baselines, checkpoints, seeds, pending updates and explicit axis sizes are not rewritten. A custom runtime must support native view operations or reports `VIEW_PREFERENCES_UNSUPPORTED`; invalid props report `INVALID_VIEW_PREFERENCES` through `onError`.

At 100%, new worksheets use 21.78 px rows and 87.12 px columns (10% larger than 19.8/79.2). The owner also approved these defaults for historical CRDT-added worksheets that have no saved default dimensions, and virtual recovery sheets. Immutable baseline worksheets keep saved defaults and omitted fields; explicit per-row/per-column sizes remain overrides. This approved default projection changes no schema-6 bytes or lineage and performs no migration/backfill. Earlier packages may display unsized appended/recovery sheets with earlier defaults; rollback must preserve later edits. Hosts creating their own initial baseline should supply the same dimensions explicitly.

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
| `@smartdoca/sheet/preferences` | Pure personal-view preference type and validator. |
| `@smartdoca/sheet/xlsx` | `xlsxToSnapshot` and `snapshotToXlsx`. |
