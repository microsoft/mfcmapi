#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
	// So all memory detached from this class is owned by a parent and must not be freed manually
	static std::wstring SUBRESCLASS = L"CResSubResEditor"; // STRING_OK
	class CResSubResEditor : public CEditor
	{
	public:
		CResSubResEditor(
			_In_ CWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			ULONG ulSubObject,
			_In_ const _SRestriction* lpRes,
			_In_ LPVOID lpAllocParent);

		void OnEditAction1() override;
		_Check_return_ LPSRestriction DetachModifiedSRestriction() noexcept;

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;

		LPVOID m_lpAllocParent;
		const _SRestriction* m_lpOldRes;
		LPSRestriction m_lpNewRes;
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
	};
} // namespace dialog::editor