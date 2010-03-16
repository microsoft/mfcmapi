#pragma once
// MFCMAPI.h : header file
//

// Version history:
// 1
//    Original, unversioned header
// 2
//    NAMEID_ARRAY_ENTRY augmented with ulType and lpszArea

#define MFCMAPI_HEADER_V1 1
#define MFCMAPI_HEADER_V2 2

// Version number of this header file. Will be bumped when breaking changes are made
#define MFCMAPI_HEADER_CURRENT_VERSION MFCMAPI_HEADER_V2

// Property tags and types - used by GetPropTags and GetPropTypes
struct NAME_ARRAY_ENTRY
{
	ULONG ulValue;
	LPWSTR lpszName;
};
typedef NAME_ARRAY_ENTRY FAR * LPNAME_ARRAY_ENTRY;

// Guids - used by GetPropGuids
struct GUID_ARRAY_ENTRY
{
	LPCGUID	lpGuid;
	LPWSTR	lpszName;
};
typedef GUID_ARRAY_ENTRY FAR * LPGUID_ARRAY_ENTRY;

// Named property mappings - used by GetNameIDs
struct NAMEID_ARRAY_ENTRY
{
	LONG    lValue;
	LPCGUID lpGuid;
	LPWSTR  lpszName;
	ULONG   ulType;
	LPWSTR  lpszArea;
};
typedef NAMEID_ARRAY_ENTRY FAR * LPNAMEID_ARRAY_ENTRY;

// Types of flag array entries
enum __FlagType
{
	flagFLAG,
	flagVALUE,
	flagVALUE3RDBYTE,
	flagVALUE4THBYTE,
	flagVALUELOWERNIBBLE,
	flagCLEARBITS, // Used to clear bits that we know we don't know so that remaining bits can be examined as values
};

// Types of guids for InterpretFlag named property lookup
enum __GuidType
{
	guidPSETID_Meeting     = 0x100,
	guidPSETID_Address     = 0x200,
	guidPSETID_Task        = 0x300,
	guidPSETID_Appointment = 0x400,
	guidPSETID_Common      = 0x500,
	guidPSETID_Log         = 0x600,
	guidPSETID_PostRss     = 0x700,
	guidPSETID_Sharing     = 0x800,
	guidPSETID_Note        = 0x900,
};

// All MAPI props are stored in the array by their PROP_ID. So all are < 0xffff.
#define FLAG_ENTRY(_fName,_fValue,_fType) {PROP_ID(_fName),(_fValue),(_fType),L#_fValue},
#define FLAG_ENTRY_NAMED(_fName,_fValue,_fValueName,_fType) {PROP_ID(_fName),(_fValue),(_fType),(_fValueName)},
#define FLAG_ENTRY3RDBYTE(_fName,_fValue,_fValType) {PROP_ID(_fName),(_fValue),flagVALUE3RDBYTE,L#_fValType L": " L#_fValue}, // STRING_OK
#define FLAG_ENTRY4THBYTE(_fName,_fValue,_fValType) {PROP_ID(_fName),(_fValue),flagVALUE4THBYTE,L#_fValType L": " L#_fValue}, // STRING_OK
#define CLEAR_BITS_ENTRY(_fName,_fValue) {PROP_ID(_fName),(_fValue),flagCLEARBITS,L""},

#define NAMEDPROP_FLAG_ENTRY(_fName,_fGuid,_fValue,_fType) {PROP_TAG((guid##_fGuid),(_fName)),(_fValue),(_fType),L#_fValue},
#define NAMEDPROP_FLAG_ENTRY_NAMED(_fName,_fGuid,_fValue,_fValueName,_fType) {PROP_TAG((guid##_fGuid),(_fName)),(_fValue),(_fType),(_fValueName)},
#define NAMEDPROP_CLEAR_BITS_ENTRY(_fName,_fGuid,_fValue) {PROP_TAG((guid##_fGuid),(_fName)),(_fValue),flagCLEARBITS,L""},

// I can put non property related flags with the high bit set (see enum __NonPropFlag)
#define NON_PROP_FLAG_ENTRY(_fName,_fValue,_fType) {(_fName),(_fValue),(_fType),L#_fValue},
#define NON_PROP_FLAG_ENTRY_NAMED(_fName,_fValue,_fValueName,_fType) {(_fName),(_fValue),(_fType),(_fValueName)},
#define NON_PROP_FLAG_ENTRYLOWERNIBBLE(_fName,_fValue,_fValType) {(_fName),(_fValue),flagVALUELOWERNIBBLE,L#_fValType L": " L#_fValue}, // STRING_OK
#define NON_PROP_CLEAR_BITS_ENTRY(_fName,_fValue) {(_fName),(_fValue),flagCLEARBITS,L""},

// Flag parsing array - used by GetPropFlags
struct FLAG_ARRAY_ENTRY
{
	ULONG	ulFlagName;
	LONG	lFlagValue;
	ULONG	ulFlagType;
	LPCWSTR lpszName;
};
typedef FLAG_ARRAY_ENTRY FAR * LPFLAG_ARRAY_ENTRY;

// Menu contexts - denote when a menu item should be present. May be combined.
// Used in _MenuItem and _AddInMenuParams.
#define MENU_CONTEXT_MAIN                 0x00000001 // Main window
#define MENU_CONTEXT_FOLDER_TREE          0x00000002 // Folder hierarchy window
#define MENU_CONTEXT_FOLDER_CONTENTS      0x00000004 // Folder contents window
#define MENU_CONTEXT_AB_TREE              0x00000008 // Address book hierarchy window
#define MENU_CONTEXT_AB_CONTENTS          0x00000010 // Address book contents window
#define MENU_CONTEXT_PROFILE_LIST         0x00000020 // Profile list window
#define MENU_CONTEXT_PROFILE_SERVICES     0x00000040 // Profile services window
#define MENU_CONTEXT_PROFILE_PROVIDERS    0x00000080 // Profile providers window
#define MENU_CONTEXT_RECIPIENT_TABLE      0x00000100 // Recipient table window
#define MENU_CONTEXT_ATTACHMENT_TABLE     0x00000200 // Attachment table window
#define MENU_CONTEXT_ACL_TABLE            0x00000400 // ACL table window
#define MENU_CONTEXT_RULES_TABLE          0x00000800 // Rules table window
#define MENU_CONTEXT_FORM_CONTAINER       0x00001000 // Form container window
#define MENU_CONTEXT_STATUS_TABLE         0x00002000 // Status table window
#define MENU_CONTEXT_RECIEVE_FOLDER_TABLE 0x00004000 // Receive folder window
#define MENU_CONTEXT_HIER_TABLE           0x00008000 // Hierarchy table window
#define MENU_CONTEXT_DEFAULT_TABLE        0x00010000 // Default table window
#define MENU_CONTEXT_MAILBOX_TABLE        0x00020000 // Mailbox table window
#define MENU_CONTEXT_PUBLIC_FOLDER_TABLE  0x00040000 // Public Folder table window
#define MENU_CONTEXT_PROPERTY             0x00080000 // Property pane

// Some helpful constants for common context scenarios
#define MENU_CONTEXT_ALL_FOLDER MENU_CONTEXT_FOLDER_TREE|MENU_CONTEXT_FOLDER_CONTENTS
#define MENU_CONTEXT_ALL_AB MENU_CONTEXT_AB_TREE|MENU_CONTEXT_AB_CONTENTS
#define MENU_CONTEXT_ALL_PROFILE MENU_CONTEXT_PROFILE_LIST|MENU_CONTEXT_PROFILE_SERVICES|MENU_CONTEXT_PROFILE_PROVIDERS
#define MENU_CONTEXT_ALL_WINDOWS \
	MENU_CONTEXT_MAIN| \
	MENU_CONTEXT_ALL_FOLDER| \
	MENU_CONTEXT_ALL_AB| \
	MENU_CONTEXT_ALL_PROFILE| \
	MENU_CONTEXT_RECIPIENT_TABLE| \
	MENU_CONTEXT_ATTACHMENT_TABLE| \
	MENU_CONTEXT_ACL_TABLE| \
	MENU_CONTEXT_RULES_TABLE| \
	MENU_CONTEXT_FORM_CONTAINER| \
	MENU_CONTEXT_STATUS_TABLE| \
	MENU_CONTEXT_RECIEVE_FOLDER_TABLE| \
	MENU_CONTEXT_HIER_TABLE| \
	MENU_CONTEXT_DEFAULT_TABLE| \
	MENU_CONTEXT_MAILBOX_TABLE| \
	MENU_CONTEXT_PUBLIC_FOLDER_TABLE
#define MENU_CONTEXT_ALL MENU_CONTEXT_ALL_WINDOWS|MENU_CONTEXT_PROPERTY

// Menu flags - denote when a menu item should be enabled, data to be fetched, etc. May be combined.
// Used in _MenuItem and _AddInMenuParams.
#define MENU_FLAGS_FOLDER_REG    0x00000001 // Enable for regular folder contents
#define MENU_FLAGS_FOLDER_ASSOC  0x00000002 // Enable for associated (FAI) folder contents
#define MENU_FLAGS_DELETED       0x00000004 // Enable for deleted items and folders
#define MENU_FLAGS_ROW           0x00000008 // Row data only
#define MENU_FLAGS_SINGLESELECT  0x00000010 // Enable only when a single item selected
#define MENU_FLAGS_MULTISELECT   0x00000020 // Enable only when multiple items selected
#define MENU_FLAGS_REQUESTMODIFY 0x00000040 // Request modify when opening item

#define MENU_FLAGS_ALL_FOLDER \
	MENU_FLAGS_FOLDER_REG| \
	MENU_FLAGS_FOLDER_ASSOC| \
	MENU_FLAGS_DELETED

struct _AddIn;
typedef _AddIn FAR * LPADDIN;

// Used by GetMenus
struct _MenuItem
{
	LPWSTR	szMenu;		// Menu string
	LPWSTR	szHelp;		// Menu help string
	ULONG	ulContext;	// Context under which menu item should be available (MENU_CONTEXT_*)
	ULONG	ulFlags;	// Flags (MENU_FLAGS_*)
	ULONG	ulID;		// Menu ID (For each add-in, each ulID must be unique for each menu item)
	LPADDIN lpAddIn;	// Pointer to AddIn structure, set by LoadAddIn - must be NULL in return from GetMenus
};
typedef _MenuItem FAR * LPMENUITEM;

// Used in CallMenu
struct _AddInMenuParams
{
	LPMENUITEM            lpAddInMenu; // the menu item which was invoked
	ULONG                 ulAddInContext; // the context in which the menu item was invoked (MENU_CONTEXT_*)
	ULONG                 ulCurrentFlags; // flags which apply (MENU_FLAGS_*)
	HWND                  hWndParent; // handle to the current window (useful for dialogs)
	// The following are the MAPI objects/data which may have been available when the menu was invoked
	LPMAPISESSION         lpMAPISession;
	LPMDB                 lpMDB;
	LPMAPIFOLDER          lpFolder;
	LPMESSAGE             lpMessage;
	LPMAPITABLE           lpTable;
	LPADRBOOK             lpAdrBook;
	LPABCONT              lpAbCont;
	LPMAILUSER            lpMailUser;
	LPEXCHANGEMODIFYTABLE lpExchTbl;
	LPMAPIFORMCONTAINER   lpFormContainer;
	LPMAPIFORMINFO        lpFormInfoProp;
	LPPROFSECT            lpProfSect;
	LPATTACH              lpAttach;
	ULONG                 ulPropTag;
	LPSRow                lpRow;
	LPMAPIPROP            lpMAPIProp; // the selected MAPI object - only used in MENU_CONTEXT_PROPERTY context
};
typedef _AddInMenuParams FAR * LPADDINMENUPARAMS;

// used in _AddInDialogControl and _AddInDialogControlResult
enum __AddInControlType
{
	ADDIN_CTRL_CHECK,
	ADDIN_CTRL_EDIT_TEXT,
	ADDIN_CTRL_EDIT_BINARY,
	ADDIN_CTRL_EDIT_NUM_DECIMAL,
	ADDIN_CTRL_EDIT_NUM_HEX
};

// used in _AddInDialog
struct _AddInDialogControl
{
	__AddInControlType cType; // Type of the control
	BOOL bReadOnly; // Whether the control should be read-only or read-write
	BOOL bMultiLine; // Whether the control should be single or multiple lines. ADDIN_CTRL_EDIT_* only.
	BOOL bDefaultCheckState; // Default state for a check box. ADDIN_CTRL_CHECK only.
	ULONG ulDefaultNum; // Default value for an edit control. ADDIN_CTRL_EDIT_NUM_DECIMAL and ADDIN_CTRL_EDIT_NUM_HEX only.
	LPWSTR szLabel; // Text label for the control. Valid for all controls.
	LPWSTR szDefaultText; // Default value for an edit control. ADDIN_CTRL_EDIT_TEXT only.
	size_t cbBin; // Count of bytes in lpBin. ADDIN_CTRL_EDIT_BINARY only.
	LPBYTE lpBin; // Array of bytes to be rendered in edit control. ADDIN_CTRL_EDIT_BINARY only.
};
typedef _AddInDialogControl FAR * LPADDINDIALOGCONTROL;

// Values for ulButtonFlags in _AddInDialog. May be combined.
#define CEDITOR_BUTTON_OK		0x00000001 // Display an OK button
#define CEDITOR_BUTTON_CANCEL	0x00000008 // Display a Cancel button

// Passed in to ComplexDialog
struct _AddInDialog
{
	LPWSTR szTitle; // Title for dialog
	LPWSTR szPrompt; // String to display at top of dialog above controls (optional)
	ULONG ulButtonFlags; // Buttons to be present at bottom of dialog
	ULONG ulNumControls; // Number of entries in lpDialogControls
	LPADDINDIALOGCONTROL lpDialogControls; // Array of _AddInDialogControl indicating which controls should be present
};
typedef _AddInDialog FAR * LPADDINDIALOG;

// Used in _AddInDialogResult
struct _AddInDialogControlResult
{
	__AddInControlType cType; // Type of the control
	BOOL bCheckState; // Value of a check box. ADDIN_CTRL_CHECK only.
	LPWSTR szText; // Value of an edit box. ADDIN_CTRL_EDIT_TEXT only.
	ULONG ulVal; // Value of an edit box. ADDIN_CTRL_EDIT_NUM_DECIMAL and ADDIN_CTRL_EDIT_NUM_HEX only.
	size_t cbBin; // Count of bytes in lpBin. ADDIN_CTRL_EDIT_BINARY only.
	LPBYTE lpBin; // Array of bytes read from edit control. ADDIN_CTRL_EDIT_BINARY only.
};
typedef _AddInDialogControlResult FAR * LPADDINDIALOGCONTROLRESULT;

// Returned by ComplexDialog. To be freed by FreeDialogResult.
struct _AddInDialogResult
{
	ULONG ulNumControls; // Number of entries in lpDialogControlResults
	LPADDINDIALOGCONTROLRESULT lpDialogControlResults; // Array of _AddInDialogControlResult with final values from the dialog
};
typedef _AddInDialogResult FAR * LPADDINDIALOGRESULT;

// Functions exported by MFCMAPI

// Function: AddInLog
// Use: Hooks in to MFCMAPI's debug logging routines.
// Notes: Uses a 4k buffer for printf parameter expansion. All string to be printed should be smaller than 4k.
//        Uses the DbgAddIn tag (0x00010000) for printing. This tag must be set in MFCMAPI to get output.
#define szAddInLog "AddInLog" // STRING_OK
typedef void (STDMETHODVCALLTYPE ADDINLOG)(
	BOOL bPrintThreadTime, // whether to print the thread and time banner.
	LPWSTR szMsg, // the message to be printed. Uses printf syntax.
	... // optional printf style parameters
	);
typedef ADDINLOG FAR *LPADDINLOG;

// Function: SimpleDialog
// Use: Hooks MFCMAPI's CEditor class to display a simple dialog.
// Notes: Uses a 4k buffer for printf parameter expansion. All string to be displayed should be smaller than 4k.
#define szSimpleDialog "SimpleDialog" // STRING_OK
typedef HRESULT (STDMETHODVCALLTYPE SIMPLEDIALOG)(
	LPWSTR szTitle, // Title for dialog
	LPWSTR szMsg, // the message to be printed. Uses printf syntax.
	... // optional printf style parameters
	);
typedef SIMPLEDIALOG FAR *LPSIMPLEDIALOG;

// Function: ComplexDialog
// Use: Hooks MFCMAPI's CEditor class to display a complex dialog. 'Complex' means the dialog has controls.
#define szComplexDialog "ComplexDialog" // STRING_OK
typedef HRESULT (_cdecl COMPLEXDIALOG)(
									   LPADDINDIALOG lpDialog, // In - pointer to a _AddInDialog structure indicating what the dialog should look like
									   LPADDINDIALOGRESULT* lppDialogResult // Out, Optional - array of _AddInDialogControlResult structures with the values of the dialog when it was closed
									                                        // Must be freed with FreeDialogResult
									   );
typedef COMPLEXDIALOG FAR *LPCOMPLEXDIALOG;

// Function: FreeDialogResult
// Use: Free lppDialogResult returned by ComplexDialog
// Notes: Failure to free lppDialogResult will result in a memory leak
#define szFreeDialogResult "FreeDialogResult" // STRING_OK
typedef void (_cdecl FREEDIALOGRESULT)(
									   LPADDINDIALOGRESULT lpDialogResult // _AddInDialogControlResult array to be freed
									   );
typedef FREEDIALOGRESULT FAR *LPFREEDIALOGRESULT;

// Function: GetMAPIModule
// Use: Get a handle to MAPI so the add-in does not have to link mapi32.lib
// Notes: MAPI may not be loaded yet. If it's not loaded and bForce isn't set, then lphModule will return a NULL pointer.
//        If MAPI isn't loaded yet and bForce is set, then MFCMAPI will load MAPI.
//        If MAPI is loaded, bForce has no effect.
//        The handle returned is NOT ref-counted. Pay close attention to return values from GetProcAddress. Do not call FreeLibrary.
#define szGetMAPIModule "GetMAPIModule" // STRING_OK
typedef void (_cdecl GETMAPIMODULE)(
									HMODULE* lphModule,
									BOOL bForce
									);
typedef GETMAPIMODULE FAR *LPGETMAPIMODULE;

// End functions exported by MFCMAPI

// Functions exported by Add-Ins

// Function: LoadAddIn
// Use: Let the add-in know we're here, get its name.
// Notes: If this function is not present, MFCMAPI won't load the add-in. This is the only required function.
#define szLoadAddIn "LoadAddIn" // STRING_OK
typedef void (STDMETHODCALLTYPE LOADADDIN)(
	LPWSTR* szTitle // Name of the add-in
	);
typedef LOADADDIN FAR *LPLOADADDIN;

// Function: UnloadAddIn
// Use: Let the add-in know we're done - it can free any resources it has allocated
#define szUnloadAddIn "UnloadAddIn" // STRING_OK
typedef void (STDMETHODCALLTYPE UNLOADADDIN)();
typedef UNLOADADDIN FAR *LPUNLOADADDIN;

// Function: GetMenus
// Use: Returns static array of menu information
#define szGetMenus "GetMenus" // STRING_OK
typedef void (STDMETHODCALLTYPE GETMENUS)(
	ULONG* lpulMenu, // Count of _MenuItem structures in lppMenu
	LPMENUITEM* lppMenu // Array of _MenuItem structures describing menu items
	);
typedef GETMENUS FAR *LPGETMENUS;

// Function: CallMenu
// Use: Calls back to AddIn with a menu choice - addin will decode and invoke
#define szCallMenu "CallMenu" // STRING_OK
typedef HRESULT (STDMETHODCALLTYPE CALLMENU)(
	LPADDINMENUPARAMS lpParams	// Everything the add-in needs to know
	);
typedef CALLMENU FAR *LPCALLMENU;

// Function: GetPropTags
// Use: Returns a static array of property names for MFCMAPI to use in decoding properties
#define szGetPropTags "GetPropTags" // STRING_OK
typedef void (STDMETHODCALLTYPE GETPROPTAGS)(
	ULONG* lpulPropTags, // Number of entries in lppPropTags
	LPNAME_ARRAY_ENTRY* lppPropTags // Array of NAME_ARRAY_ENTRY structures
	);
typedef GETPROPTAGS FAR *LPGETPROPTAGS;

// Function: GetPropTypes
// Use: Returns a static array of property types for MFCMAPI to use in decoding properties
#define szGetPropTypes "GetPropTypes" // STRING_OK
typedef void (STDMETHODCALLTYPE GETPROPTYPES)(
	ULONG* lpulPropTypes, // Number of entries in lppPropTypes
	LPNAME_ARRAY_ENTRY* lppPropTypes // Array of NAME_ARRAY_ENTRY structures
	);
typedef GETPROPTYPES FAR *LPGETPROPTYPES;

// Function: GetPropGuids
// Use: Returns a static array of property guids for MFCMAPI to use in decoding properties
#define szGetPropGuids "GetPropGuids" // STRING_OK
typedef void (STDMETHODCALLTYPE GETPROPGUIDS)(
	ULONG* lpulPropGuids, // Number of entries in lppPropGuids
	LPGUID_ARRAY_ENTRY* lppPropGuids // Array of GUID_ARRAY_ENTRY structures
	);
typedef GETPROPGUIDS FAR *LPGETPROPGUIDS;

// Function: GetNameIDs
// Use: Returns a static array of named property mappings for MFCMAPI to use in decoding properties
#define szGetNameIDs "GetNameIDs" // STRING_OK
typedef void (STDMETHODCALLTYPE GETNAMEIDS)(
	ULONG* lpulNameIDs, // Number of entries in lppNameIDs
	LPNAMEID_ARRAY_ENTRY* lppNameIDs // Array of NAMEID_ARRAY_ENTRY structures
	);
typedef GETNAMEIDS FAR *LPGETNAMEIDS;

// Function: GetPropFlags
// Use: Returns a static array of flag parsing information for MFCMAPI to use in decoding properties
#define szGetPropFlags "GetPropFlags" // STRING_OK
typedef void (STDMETHODCALLTYPE GETPROPFLAGS)(
	ULONG* lpulPropFlags, // Number of entries in lppPropFlags
	LPFLAG_ARRAY_ENTRY* lppPropFlags // Array of FLAG_ARRAY_ENTRY structures
	);
typedef GETPROPFLAGS FAR *LPGETPROPFLAGS;

// Function: GetAPIVersion
// Use: Returns version number of the API used by the add-in
// Notes: MUST return MFCMAPI_HEADER_CURRENT_VERSION
#define szGetAPIVersion "GetAPIVersion" // STRING_OK
typedef ULONG (STDMETHODCALLTYPE GETAPIVERSION)();
typedef GETAPIVERSION FAR *LPGETAPIVERSION;

// Structure used internally by MFCMAPI to track information on loaded Add-Ins. While it is accessible
// by the add-in through the CallMenu function, it should only be consulted for debugging purposes.
struct _AddIn
{
	LPADDIN              lpNextAddIn; // Pointer to the next add-in
	HMODULE              hMod;        // Handle to add-in module
	LPWSTR               szName;      // Name of add-in
	LPCALLMENU           pfnCallMenu; // Pointer to function in module for invoking menus
	ULONG                ulMenu;      // Count of menu items exposed by add-in
	LPMENUITEM           lpMenu;      // Array of menu items exposed by add-in
	ULONG                ulPropTags;  // Count of property tags exposed by add-in
	LPNAME_ARRAY_ENTRY   lpPropTags;  // Array of property tags exposed by add-in
	ULONG                ulPropTypes; // Count of property tags exposed by add-in
	LPNAME_ARRAY_ENTRY   lpPropTypes; // Array of property tags exposed by add-in
	ULONG                ulPropGuids; // Count of property guids exposed by add-in
	LPGUID_ARRAY_ENTRY   lpPropGuids; // Array of property guids exposed by add-in
	ULONG                ulNameIDs;   // Count of named property mappings exposed by add-in
	LPNAMEID_ARRAY_ENTRY lpNameIDs;   // Array of named property mappings exposed by add-in
	ULONG                ulPropFlags; // Count of flags exposed by add-in
	LPFLAG_ARRAY_ENTRY   lpPropFlags; // Array of flags exposed by add-in
};

// Everything below this point is internal to MFCMAPI and should be removed from this header when including it in an add-in

EXTERN_C __declspec(dllexport) void __cdecl AddInLog(BOOL bPrintThreadTime, LPWSTR szMsg,...);
EXTERN_C __declspec(dllexport) HRESULT __cdecl SimpleDialog(LPWSTR szTitle, LPWSTR szMsg,...);
EXTERN_C __declspec(dllexport) HRESULT __cdecl ComplexDialog(LPADDINDIALOG lpDialog, LPADDINDIALOGRESULT* lppDialogResult);
EXTERN_C __declspec(dllexport) void __cdecl FreeDialogResult(LPADDINDIALOGRESULT lpDialogResult);
EXTERN_C __declspec(dllexport) void __cdecl GetMAPIModule(HMODULE* lphModule, BOOL bForce);

void LoadAddIns();
void UnloadAddIns();
ULONG ExtendAddInMenu(HMENU hMenu, ULONG ulAddInContext);
LPMENUITEM GetAddinMenuItem(HWND hWnd, UINT uidMsg);
void InvokeAddInMenu(LPADDINMENUPARAMS lpParams);
void MergeAddInArrays();
void SortAddInArrays();
LPNAMEID_ARRAY_ENTRY GetDispIDFromName(LPCWSTR lpszDispIDName);

extern LPNAME_ARRAY_ENTRY PropTagArray;
extern ULONG ulPropTagArray;

extern LPNAME_ARRAY_ENTRY PropTypeArray;
extern ULONG ulPropTypeArray;

extern LPGUID_ARRAY_ENTRY PropGuidArray;
extern ULONG ulPropGuidArray;

extern LPNAMEID_ARRAY_ENTRY NameIDArray;
extern ULONG ulNameIDArray;

extern LPFLAG_ARRAY_ENTRY FlagArray;
extern ULONG ulFlagArray;