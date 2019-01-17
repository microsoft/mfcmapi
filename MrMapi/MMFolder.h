#pragma once
// Folder related utilities for MrMAPI

namespace cli {
	struct MYOPTIONS;
}

void DoChildFolders(_In_ cli::MYOPTIONS ProgOpts);
void DoFolderProps(_In_ cli::MYOPTIONS ProgOpts);
void DoFolderSize(_In_ cli::MYOPTIONS ProgOpts);
void DoSearchState(_In_ cli::MYOPTIONS ProgOpts);

LPMAPIFOLDER MAPIOpenFolderExW(
	_In_ LPMDB lpMdb, // Open message store
	_In_z_ const std::wstring& lpszFolderPath); // folder path