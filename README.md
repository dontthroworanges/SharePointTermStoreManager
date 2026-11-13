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
