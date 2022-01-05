#pragma once
#include <UI/Dialogs/Editors/Editor.h>
#include <core/mapi/cache/mapiObjects.h>

namespace dialog::editor
{
	class CriteriaEditor : public CEditor
	{
	public:
		CriteriaEditor(
			_In_ CWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ const _SRestriction* lpRes,
			_In_ LPENTRYLIST lpEntryList,
			ULONG ulSearchState);
		~CriteriaEditor();

		_Check_return_ LPSRestriction DetachModifiedSRestriction() noexcept;
		_Check_return_ LPENTRYLIST DetachModifiedEntryList() noexcept;
		_Check_return_ ULONG GetSearchFlags() const noexcept;
		_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData) override;

	private:
		void OnEditAction1() override;
		BOOL OnInitDialog() override;
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void InitListFromEntryList(ULONG ulListNum, _In_ const SBinaryArray* lpEntryList) const;
		void OnOK() override;

		_Check_return_ const _SRestriction* GetSourceRes() const noexcept;

		const _SRestriction* m_lpSourceRes;
		LPSRestriction m_lpNewRes;

		LPENTRYLIST m_lpSourceEntryList;
		LPENTRYLIST m_lpNewEntryList{};

		ULONG m_ulNewSearchFlags;
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
	};
} // namespace dialog::editor