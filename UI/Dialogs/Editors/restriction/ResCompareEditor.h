#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	static std::wstring COMPCLASS = L"ResCompareEditor"; // STRING_OK
	class ResCompareEditor : public CEditor
	{
	public:
		ResCompareEditor(_In_ CWnd* pParentWnd, ULONG ulRelop, ULONG ulPropTag1, ULONG ulPropTag2);

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;
	};
} // namespace dialog::editor