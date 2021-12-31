#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	static std::wstring COMPCLASS = L"CResCompareEditor"; // STRING_OK
	class CResCompareEditor : public CEditor
	{
	public:
		CResCompareEditor(_In_ CWnd* pParentWnd, ULONG ulRelop, ULONG ulPropTag1, ULONG ulPropTag2);

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;
	};
} // namespace dialog::editor