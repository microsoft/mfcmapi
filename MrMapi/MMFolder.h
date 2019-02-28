#pragma once
// Folder related utilities for MrMAPI

void DoChildFolders(_In_ LPMAPIFOLDER lpFolder);
void DoFolderProps(LPMAPIFOLDER lpFolder);
void DoFolderSize(_In_ LPMAPIFOLDER lpFolder);
void DoSearchState(_In_ LPMAPIFOLDER lpFolder);

LPMAPIFOLDER MAPIOpenFolderExW(
	_In_ LPMDB lpMdb, // Open message store
	_In_z_ const std::wstring& lpszFolderPath); // folder path