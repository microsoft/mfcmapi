#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
	// So all memory detached from this class is owned by a parent and must not be freed manually
	static std::wstring ANDORCLASS = L"ResAndOrEditor"; // STRING_OK
	class ResAndOrEditor : public CEditor
	{
	public:
		ResAndOrEditor(
			_In_ CWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ const _SRestriction* lpRes,
			_In_ LPVOID lpAllocParent);

		_Check_return_ LPSRestriction DetachModifiedSRestrictionArray() noexcept;
		_Check_return_ ULONG GetResCount() const noexcept;
		_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData) override;

	private:
		BOOL OnInitDialog() override;
		void InitListFromRestriction(ULONG ulListNum, _In_ const _SRestriction* lpRes) const;
		void OnOK() override;

		LPVOID m_lpAllocParent;
		const _SRestriction* m_lpRes;
		LPSRestriction m_lpNewResArray;
		ULONG m_ulNewResCount;
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
	};
} // namespace dialog::editor