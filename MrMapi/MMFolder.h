#pragma once
// Folder related utilities for MrMAPI

void DoChildFolders(_In_opt_ LPMAPIFOLDER lpFolder);
void DoFolderProps(_In_opt_ LPMAPIFOLDER lpFolder);
void DoFolderSize(_In_opt_ LPMAPIFOLDER lpFolder);
void DoSearchState(_In_opt_ LPMAPIFOLDER lpFolder);

LPMAPIFOLDER MAPIOpenFolderExW(
	_In_opt_ LPMDB lpMdb, // Open message store
	_In_ const std::wstring& lpszFolderPath); // folder path