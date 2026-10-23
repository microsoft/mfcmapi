Add-ins are disabled by default in both MFCMAPI and MrMAPI. To enable them, build a DLL that exports the functions below and place it in a directory on a local fixed drive. In MFCMAPI, open **Options**, find **Add-ins**, and enter or use **Browse...** to select the **Add-in directory**. Canceling the picker leaves the current field unchanged; the setting is saved only when Options is accepted. Restart MFCMAPI to apply it. MrMAPI uses the same per-user setting on its next invocation; no directory switch is needed.

The directory must be an existing, accessible, absolute local path, such as `C:\Tools\MFCMAPIAddins`. Relative paths, network paths (including mapped network drives), removable drives, and environment-variable expansion are not supported. Only DLLs directly in that directory are scanned, not subdirectories. Neither application searches beside its executable unless you explicitly configure that directory. Clear the field and accept Options to disable add-ins on subsequent launches; currently loaded add-ins remain loaded until exit.

The setting is stored as `AddInDirectory` (`REG_SZ`) under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\MFCMAPI`. Missing or empty means disabled. Invalid or unavailable directories produce a diagnostic and never fall back to another location. The old `LoadAddIns` setting is ignored: after upgrading, explicitly configure a directory even if you previously enabled add-ins.

On startup, MFCMAPI checks DLL exports without loading the file and offers add-in-shaped DLLs that have not been seen before in one approval dialog. Check the add-ins you want to load. Configuring a directory does not approve its contents.

MFCMAPI remembers both approvals and denials by the DLL's SHA-256 hash. Replacing an add-in with different contents requires a new decision; renaming an identical DLL within the configured directory does not. Unchecked entries are remembered as denied. Canceling the approval dialog saves no decisions for the new entries.

Changing or clearing the add-in directory and accepting Options automatically clears all remembered approvals and denials. Leaving the directory unchanged, canceling Options, or entering an invalid directory does not clear them. Previously approved add-ins require new approval on the next launch after a directory change; currently loaded add-ins remain loaded until exit. You can also use the **Clear** action in Advanced Options to clear all remembered decisions without changing the directory. To disable a previously approved add-in, clear its approval and deny it, or rename it so it no longer has a `.dll` extension.

With a directory configured, MrMAPI loads previously approved add-ins without prompting. It skips unknown add-ins by default; pass `-ApproveAddins` to enable interactive yes/no prompts. This switch does not enable add-ins without a configured directory. `-NoAddins` bypasses discovery and loading for that invocation, takes precedence over `-ApproveAddins`, and does not change the saved setting.

MrMAPI's registry option list contains only `AddInDirectory`; other saved MFCMAPI options do not change its console defaults or command-line behavior. The add-in loader uses the initialized option directly; it does not perform a separate registry read.

In MFCMAPI, **Tools > Options > Advanced > Add-ins** includes a read-only, scrollable discovery report alongside the approval controls. The directory setting remains on the main Options page. When opening Advanced, the report shows the directory currently entered in Options separately from the directory used at startup, and lists each DLL candidate with its recorded outcome: loaded, awaiting approval, denied, incompatible, unreadable, or failed to load. A candidate's presence does not mean it was loaded. Windows errors are shown when available. Reopen Advanced after editing the directory to see the pending path; this does not save the setting or change the recorded discovery results. Opening Advanced does not rescan, load DLLs, prompt for approvals, or change decisions; restart after accepting a directory change to run discovery again. Cancelling Options leaves the saved directory unchanged.

The core of an MFCMAPI addin is the header file: [MFCMAPI.h](MFCMAPI.h). This file contains the definitions for the functions you'll need to implement and export from your DLL, as well as definitions for functions your add-in can import from MFCMAPI.

To get you started, here's a test add-in: [TestAddIn.zip](TestAddIn.zip). It doesn't do much except exercise the various add-in entry points. It was developed in sync with the addition of add-ins to the code base as part of my testing.

# Documentation
## Stuff the Add-In implements
### LoadAddIn
This function lets the add-in know we're here and get its name. You have to implement this one for MFCMAPI to know you're an add-in.

### UnloadAddIn
MFCMAPI will call this when unloading add-ins. This is a good time to free resources.

### GetMenus
Return an array describing the menu entries you'd like to have, when you want them to be enabled, and what information you want passed back to you when your menu is invoked.

### CallMenu
Let's you know one of your menu items has been selected, along with a wealth of contextual information.

### GetPropTags
Return an array of property names for MFCMAPI to use in decoding properties.

### GetPropTypes
Return an array of property types for MFCMAPI to use in decoding properties.

### GetPropGuids
Return an array of property guids for MFCMAPI to use in decoding properties.

### GetNameIDs
Returns an array of named property mappings for MFCMAPI to use in decoding properties.

### GetPropFlags
Return an array of flag parsing information for MFCMAPI to use in decoding properties.

## Stuff MFCMAPI implements
### AddInLog
Call this to generate debug output through MFCMAPI's own logging routines.

### SimpleDialog
A quick and dirty dialog.

### ComplexDialog
A harder to use, but much more featured dialog that can be used to get user input.

### FreeDialogResult
If you use ComplexDialog, you might get a structure back with the results. This function frees that structure.

### GetMAPIModule
Get a handle to the MAPI binary that MFCMAPI loaded, or ask MFCMAPI to go ahead and load MAPI.

## Notes
A well behaved add-in will not link mapi32.lib. Instead, it will use GetMAPIModule to get a handle to MAPI, then [GetProcAddress](http://msdn2.microsoft.com/en-us/library/ms683212.aspx) to get any exports you need. The reason for not statically linking is to allow MFCMAPI the flexibility to load MAPI itself from custom locations.

Also, a well behaved add-in is not written in .Net, since MAPI and .Net aren't supported in combination. No bugs will be taken on this.
