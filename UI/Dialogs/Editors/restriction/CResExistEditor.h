#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	static std::wstring EXISTCLASS = L"CResExistEditor"; // STRING_OK
	class CResExistEditor : public CEditor
	{
	public:
		CResExistEditor(_In_ CWnd* pParentWnd, ULONG ulPropTag);

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;
	};
} // namespace dialog::editor