#pragma once

namespace dialog::editor
{
	// Returns true if options changed that may require a property refresh
	bool DisplayOptionsDlg(_In_ CWnd* lpParentWnd);
} // namespace dialog::editor