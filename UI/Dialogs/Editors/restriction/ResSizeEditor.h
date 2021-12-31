#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	static std::wstring SIZECLASS = L"ResSizeEditor"; // STRING_OK
	class ResSizeEditor : public CEditor
	{
	public:
		ResSizeEditor(_In_ CWnd* pParentWnd, ULONG relop, ULONG ulPropTag, ULONG cb);

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;
	};
} // namespace dialog::editor