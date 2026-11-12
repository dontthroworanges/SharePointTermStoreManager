# SharePoint Term Store Manager
A WinUI 3 desktop app for managing the SharePoint Online term store, built to be friendlier than the term store
admin page in the SharePoint admin center.

> [!IMPORTANT]
> **Disclaimer:** This application is being built with AI-generated code. If you prefer not to use AI-generated software, please do not proceed.

## Status
Early development. Today the app can:
- Save one or more tenant profiles and sign in (no app registration required). The navigation pane on the left,
  as in SharePoint Search Explorer, holds the term store, a profile switcher (switch site, connect or refresh,
  change account, edit, sign out, remove) and Settings (Connections, Options, About).
- Open the default term store and browse groups, term sets and terms in a tree (children load on expand,
  in the term set's custom sort order).
- Resize the tree by dragging the divider between the tree and the editor (or focus it and use the arrow keys);
  double-click it to go back to the default split. The width is remembered.
- Create term groups (**New term group** above the tree), term sets and terms (the command bar above the
  selected item, or right-click an item).
- Edit the selected item in collapsible sections laid out like the SharePoint term store's tabs: **General**
  (default label or name, synonyms, description, term set owner and contact; other languages' labels read-only),
  **Usage settings** (available for tagging, submission policy, sort order, unique identifier, the term sets a term
  is a member of, and details) and **Advanced** (shared and local custom properties). Labels and descriptions are edited
  in the term store's default language; labels in other languages are left as they are.
- Rename (F2) and delete (Delete key), or use the command bar or right-click. Term groups must be empty to be
  deleted. System groups and their term sets can't be renamed or deleted, and pinned copies of terms can only be
  changed at their source.
- Before saving, the app checks that nobody changed the item since it was opened, so it never overwrites someone
  else's edit. A reused term is changed and deleted only in the term set it's selected in.
- Term actions (the command bar above the selected term, or right-click it): move, copy, reuse, pin, merge, deprecate and
  un-deprecate. Term sets can be moved to another group.
- **Merge** and the other actions use a target picker: search labels and synonyms (starts with) in one term
  set or the whole term store, or browse. Merge shows a preview of the labels the target gains before anything
  changes. Terms with child terms aren't merged; move or merge the children first.
- Search the whole term store from the box in the title bar: terms whose label or synonym starts with the text,
  term sets named exactly the text (or loaded ones containing it), and term groups containing it. Choosing a
  result opens the tree down to it and selects it.
- Deprecated terms are dimmed and shown with a "tag off" icon.
- Settings › Appearance: app theme (system, light or dark), accent colour (Windows accent or one of the Windows
  colours) and outline or colourful term store icons (coloured groups, term sets, and reused, pinned and
  deprecated terms; plain terms keep their outline icon), as in SharePoint Search Explorer.
- Settings › Options: how many seconds success messages stay before closing themselves (0 keeps
  them open). Errors and warnings stay until closed.
- An operations log (Settings › Options › Open folder) records every change as JSON Lines in
  `%LOCALAPPDATA%\SharePointTermStoreManager\Logs`, with a copy of each item's details before it's merged or
  deleted (not including child terms).
- The Operations Log page (below Term Store in the navigation) shows the log newest first, with a search box
  and filters for date, action, outcome, term group and site. Each entry expands to its JSON, indented with
  guide lines. Copy the entries shown (as JSON Lines or formatted JSON), or open a month's raw `.jsonl` file.
- The Export page (below Operations Log) saves the term store, read fresh from SharePoint, to a file: an Excel
  workbook or CSV (one row per term group, term set and term, in an Excel table), pretty-printed JSON nested
  like the tree, or a single self-contained HTML page with a collapsible tree, search and a details pane.
  Choose the term groups and whether to include deprecated terms. The Excel file is written directly as Office
  Open XML, so no extra library is needed.

## Planned
- A bulk merge queue (multi-select or CSV of source → target), with validation, preview, progress and a log.
- Snapshot export of whole term sets (with child terms) before destructive operations.

## How it talks to SharePoint
The app uses SharePoint CSOM (`Microsoft.SharePoint.Client.Taxonomy`). It is the only supported API that covers
CRUD, merge, reuse and deprecate without an app registration.

Sign-in uses Microsoft's public SharePoint Online Management Shell client, the same approach the
SharePoint Search Explorer app and SharePoint's own PowerShell module use, so no app
registration is needed. Microsoft doesn't support using that client from other apps, and could restrict it in
future. If your organization blocks it, or provides its own app registration, enter that app's client ID under
**Advanced** in the profile.

Your account needs term store permissions (term store admin, or group manager/contributor) to make changes.
SharePoint Online only; SharePoint Server (on-premises) is not supported.

## Build and run
Requirements: Windows 10 1809 or later, the .NET 8 SDK (or later), and the Microsoft Edge WebView2 Runtime
(installed with Windows 11).

```
dotnet build SharePointTermStoreManager/SharePointTermStoreManager.csproj
```

The build output runs as is; there is no installer. Settings, the sign-in session and the encrypted token cache are
kept in `%LOCALAPPDATA%\SharePointTermStoreManager`. **Settings › About › Reset app to defaults** signs out and
removes every profile, clears the sign-in session and restores all settings; the operations log is kept unless you
choose to delete it too.
