#pragma once
// Folder related utilities for MrMAPI

void DoChildFolders(_In_ MYOPTIONS ProgOpts);
void DoFolderProps(_In_ MYOPTIONS ProgOpts);
void DoFolderSize(_In_ MYOPTIONS ProgOpts);
void DoSearchState(_In_ MYOPTIONS ProgOpts);

LPMAPIFOLDER MAPIOpenFolderExW(
	_In_ LPMDB lpMdb, // Open message store
	_In_z_ const std::wstring& lpszFolderPath); // folder path