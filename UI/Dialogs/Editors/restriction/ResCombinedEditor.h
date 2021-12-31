#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	// This class is only invoked by RestrictEditor. RestrictEditor always passes an alloc parent.
	// So all memory detached from this class is owned by a parent and must not be freed manually
	static std::wstring CONTENTCLASS = L"ResCombinedEditor"; // STRING_OK
	class ResCombinedEditor : public CEditor
	{
	public:
		ResCombinedEditor(
			_In_ CWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			ULONG ulResType,
			ULONG ulCompare,
			ULONG ulPropTag,
			_In_ const _SPropValue* lpProp,
			_In_ LPVOID lpAllocParent);

		void OnEditAction1() override;
		_Check_return_ LPSPropValue DetachModifiedSPropValue() noexcept;

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;

		ULONG m_ulResType;
		LPVOID m_lpAllocParent;
		const _SPropValue* m_lpOldProp;
		LPSPropValue m_lpNewProp;
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
	};
} // namespace dialog::editor