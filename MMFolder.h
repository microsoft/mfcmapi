#pragma once
// MMFolder.h : Folder related utilities for MrMAPI

void DoChildFolders(_In_ MYOPTIONS ProgOpts);

STDMETHODIMP OpenDefaultFolder(_In_ ULONG ulFolder, _In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER *lpFolder);
HRESULT HrMAPIOpenFolderExW(
	_In_ LPMDB lpMdb,                        // Open message store
	_In_z_ LPCWSTR lpszFolderPath,           // folder path
	_Deref_out_opt_ LPMAPIFOLDER* lppFolder);// pointer to folder opened

enum
{
	DEFAULT_UNSPECIFIED,
	DEFAULT_CALENDAR,
	DEFAULT_CONTACTS,
	DEFAULT_JOURNAL,
	DEFAULT_NOTES,
	DEFAULT_TASKS,
	DEFAULT_REMINDERS,
	DEFAULT_DRAFTS,
	DEFAULT_SENTITEMS,
	DEFAULT_OUTBOX,
	DEFAULT_DELETEDITEMS,
	DEFAULT_FINDER,
	DEFAULT_IPM_SUBTREE,
	DEFAULT_INBOX,
	DEFAULT_LOCALFREEBUSY,
	DEFAULT_CONFLICTS,
	DEFAULT_SYNCISSUES,
	DEFAULT_LOCALFAILURES,
	DEFAULT_SERVERFAILURES,
	DEFAULT_JUNKMAIL,
	NUM_DEFAULT_PROPS
};

// Keep this in sync with the above enum
static LPSTR FolderNames[] = {
	"",
	"Calendar",
	"Contacts",
	"Journal",
	"Notes",
	"Tasks",
	"Reminders",
	"Drafts",
	"Sent Items",
	"Outbox",
	"Deleted Items",
	"Finder",
	"IPM_SUBTREE",
	"Inbox",
	"Local Freebusy",
	"Conflicts",
	"Sync Issues",
	"Local Failures",
	"Server Failures",
	"Junk E-mail",
};