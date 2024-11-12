# PinnedApps

A PowerShell module for pinning and un-pinning apps from the Windows start menu and taskbar.

The content I have here is an early concept that does not yet pin/unpin apps, but is able to retrieve the list (albeit slowly).

Update: I learned that the old InvokeVerb() method was revoked for start menu items in Windows 10 build 1903. It still works for taskbar items, but for the start menu, GPO/MDM and Export/Import-StartLayout are the only options. 😝
