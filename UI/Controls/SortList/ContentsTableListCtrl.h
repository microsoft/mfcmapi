#pragma once
#include <UI/Controls/SortList/SortListCtrl.h>
#include <UI/enums.h>
#include <core/mapi/columnTags.h>

namespace cache
{
	class CMapiObjects;
}

namespace dialog
{
	class CContentsTableDlg;
}

namespace mapi
{
	class adviseSink;
} // namespace mapi

namespace controls::sortlistctrl
{
	class CContentsTableListCtrl : public CSortListCtrl
	{
	public:
		const static ULONG NUMROWSPERLOOP{255};

		CContentsTableListCtrl(
			_In_ CWnd* pCreateParent,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ LPSPropTagArray sptExtraColumnTags,
			_In_ const std::vector<columns::TagNames>& lpExtraDisplayColumns,
			UINT nIDContextMenu,
			bool bIsAB,
			_In_ dialog::CContentsTableDlg* lpHostDlg);
		~CContentsTableListCtrl();

		// Initialization
		void
		SetContentsTable(_In_opt_ LPMAPITABLE lpContentsTable, tableDisplayFlags displayFlags, ULONG ulContainerType);

		// Selected item accessors
		void CopyRows() const;
		_Check_return_ LPENTRYLIST GetSelectedItemEIDs() const;
		_Check_return_ sortlistdata::sortListData* GetSortListData(int iItem) const;
		_Check_return_ LPMAPIPROP OpenNextSelectedItemProp(_Inout_opt_ int* iCurItem, modifyType bModify) const;
		_Check_return_ std::vector<int> GetSelectedItemNums() const;
		_Check_return_ std::vector<sortlistdata::sortListData*> GetSelectedItemData() const;
		_Check_return_ sortlistdata::sortListData* GetFirstSelectedItemData() const;

		_Check_return_ HRESULT ApplyRestriction() const;
		_Check_return_ LPMAPIPROP DefaultOpenItemProp(int iItem, modifyType bModify) const;
		void NotificationOff();
		void NotificationOn();
		void RefreshTable();
		void OnCancelTableLoad();
		void OnOutputTable(const std::wstring& szFileName) const;
		void SetSortTable(_In_ LPSSortOrderSet lpSortOrderSet, ULONG ulFlags) const;
		void SetUIColumns(_In_ LPSPropTagArray lpTags);
		_Check_return_ bool IsLoading() const noexcept;
		void ClearLoading() noexcept;
		void SetRestriction(_In_opt_ const _SRestriction* lpRes) noexcept;
		_Check_return_ const _SRestriction* GetRestriction() const noexcept;
		_Check_return_ restrictionType GetRestrictionType() const noexcept;
		void SetRestrictionType(restrictionType RestrictionType) noexcept;
		_Check_return_ ULONG GetContainerType() const noexcept;
		_Check_return_ bool IsAdviseSet() const noexcept;
		_Check_return_ bool IsContentsTableSet() const noexcept;
		void DoSetColumns(bool bAddExtras, bool bDisplayEditor);
		void GetStatus();
		bool bAbortLoad() noexcept { return m_bAbortLoad != 0; }

	private:
		const static ULONG NODISPLAYNAME{0xffffffff};

		// Overrides from base class
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
		void OnItemChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		void OnContextMenu(_In_ CWnd* pWnd, CPoint pos);

		void AddColumn(UINT uidHeaderName, ULONG ulCurHeaderCol, ULONG ulCurTagArrayRow, ULONG ulPropTag);
		void AddColumns(_In_ LPSPropTagArray lpCurColTagArray);
		void AddItemToListBox(int iRow, _In_ LPSRow lpsRowToAdd);
		_Check_return_ HRESULT DoExpandCollapse();
		_Check_return_ int FindRow(_In_ const SBinary& instance) const;
		_Check_return_ int GetNextSelectedItemNum(_Inout_opt_ int* iCurItem) const;
		void LoadContentsTableIntoView();
		void RefreshItem(int iRow, _In_ LPSRow lpsRowData, bool bItemExists);
		void SelectAll();
		void SetRowStrings(int iRow, _In_ LPSRow lpsRowData);

		// Custom messages
		_Check_return_ LRESULT msgOnAddItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnModifyItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnRefreshTable(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnThreadAddItem(WPARAM wParam, LPARAM lParam);

		tableDisplayFlags m_displayFlags{tableDisplayFlags::dfNormal};
		LONG volatile m_bAbortLoad{};
		std::thread m_LoadThreadHandle{};
		dialog::CContentsTableDlg* m_lpHostDlg{};
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
		std::vector<columns::TagNames> m_lpDefaultDisplayColumns{};
		LPSPropTagArray m_sptDefaultDisplayColumnTags{};
		ULONG m_ulHeaderColumns{};
		ULONG_PTR m_ulAdviseConnection{};
		ULONG m_ulDisplayNameColumn{NODISPLAYNAME};
		UINT m_nIDContextMenu{};
		bool m_bIsAB{};
		bool m_bInLoadOp{};
		const _SRestriction* m_lpRes{};
		ULONG m_ulContainerType{};
		mapi::adviseSink* m_lpAdviseSink{};
		LPMAPITABLE m_lpContentsTable{};

		restrictionType m_RestrictionType{restrictionType::none};

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls::sortlistctrl