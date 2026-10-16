The model for add-in development for MFCMAPI is quite simple. Build a DLL that exports a few functions and drop it into the same directory as MFCMAPI. On startup, MFCMAPI checks DLL exports without loading the file and offers add-in-shaped DLLs that have not been seen before in one approval dialog. Check the add-ins you want to load.

MFCMAPI remembers both approvals and denials by the DLL's SHA-256 hash. Replacing an add-in with different contents requires a new decision; moving an identical DLL does not. Unchecked entries are remembered as denied. Canceling the dialog saves no decisions for the new entries. Use the **Clear** action in the Options dialog to clear all remembered add-in decisions. To disable a previously approved add-in, clear its approval and deny it, or rename it so it no longer has a `.dll` extension.

MrMAPI loads previously approved add-ins without prompting. It skips unknown add-ins by default; pass `-ApproveAddins` to enable interactive yes/no prompts. `-NoAddins` disables add-in loading.

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
