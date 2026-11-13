# SharePoint Term Store Manager
A WinUI 3 desktop app for managing the SharePoint Online term store, built to be friendlier than the term store
admin page in the SharePoint admin center.

> [!IMPORTANT]
> **Disclaimer:** This application is being built with AI-generated code. If you prefer not to use AI-generated software, please do not proceed.

## About
I got fed up with not having the ability to search for terms when doing merge actions... So here we are. Anyway, this desktop application replicates all of the functionality of the Web based term store manager but with some additional improvements and tools. Aside from the afformentioned search when merging, the app can also export selected or all term groups into multiple formats. Finally, since I have been working with Power Automate a lot lately, there is also an easily copied term|guid string for use in different actions. I'll probably add more features to the application as needed. Okay byeeeeeee!

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

## Updates
**Settings › About › Check for updates** (and, unless turned off, every start) asks GitHub for the latest release of
[dontthroworanges/SharePointTermStoreManager](https://github.com/dontthroworanges/SharePointTermStoreManager). If its
tag (e.g. `v0.2.0`) is newer than the app's `<Version>`, the release's `.exe` is downloaded in the background to
`%LOCALAPPDATA%\SharePointTermStoreManager\Updates`, checked against GitHub's SHA-256 digest, and offered for install.
Installing renames the running `.exe` to `.exe.old`, puts the new one in its place and restarts; the next start removes
the `.old` file. If the app's folder needs admin rights, Windows asks for them.

To publish a release, bump `<Version>` in the csproj, then publish the single-file `.exe` (also the VS Code `publish`
task):

```
dotnet publish SharePointTermStoreManager/SharePointTermStoreManager.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -o publish/win-x64
```

This produces `publish/win-x64/SharePointTermStoreManager.exe`. Attach it to a GitHub release tagged `v<Version>`
(drafts and pre-releases are ignored), keeping its file name: an update saves the new `.exe` under the release file's
name, in place of the running one. If you publish more than one architecture, include `x64`, `x86` or `arm64` in each
asset's name.

## Install options
The app is a single portable `.exe` that runs from any folder, with nothing to install. To add it to the Start menu,
turn on **Settings › About › Show in the Start menu**. This creates a shortcut for your account only, without admin
rights. The app keeps the shortcut pointing at itself after it updates or moves. Don't rename the `.exe`: a WinUI app
finds its resources by its own file name.
