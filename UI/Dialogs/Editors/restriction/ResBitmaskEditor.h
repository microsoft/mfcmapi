#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	static std::wstring BITMASKCLASS = L"ResBitmaskEditor"; // STRING_OK
	class ResBitmaskEditor : public CEditor
	{
	public:
		ResBitmaskEditor(_In_ CWnd* pParentWnd, ULONG relBMR, ULONG ulPropTag, ULONG ulMask);

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;
	};
} // namespace dialog::editor