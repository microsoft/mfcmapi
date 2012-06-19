#pragma once
// MMFolder.h : Folder related utilities for MrMAPI

void DoChildFolders(_In_ MYOPTIONS ProgOpts);
void DoFolderProps(_In_ MYOPTIONS ProgOpts);

HRESULT HrMAPIOpenFolderExW(
	_In_ LPMDB lpMdb,                        // Open message store
	_In_z_ LPCWSTR lpszFolderPath,           // folder path
	_Deref_out_opt_ LPMAPIFOLDER* lppFolder);// pointer to folder opened

// Keep this in sync with the NUM_DEFAULT_PROPS enum in MAPIFunctions.h
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