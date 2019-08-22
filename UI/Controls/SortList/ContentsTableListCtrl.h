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

namespace controls
{
	namespace sortlistctrl
	{
		class CContentsTableListCtrl : public CSortListCtrl
		{
		public:
			const static ULONG NUMROWSPERLOOP{255};

			CContentsTableListCtrl(
				_In_ CWnd* pCreateParent,
				_In_ cache::CMapiObjects* lpMapiObjects,
				_In_ LPSPropTagArray sptExtraColumnTags,
				_In_ const std::vector<columns::TagNames>& lpExtraDisplayColumns,
				UINT nIDContextMenu,
				bool bIsAB,
				_In_ dialog::CContentsTableDlg* lpHostDlg);
			virtual ~CContentsTableListCtrl();

			// Initialization
			void SetContentsTable(_In_opt_ LPMAPITABLE lpContentsTable, ULONG ulDisplayFlags, ULONG ulContainerType);

			// Selected item accessors
			_Check_return_ LPENTRYLIST GetSelectedItemEIDs() const;
			_Check_return_ sortlistdata::sortListData* GetSortListData(int iItem) const;
			_Check_return_ LPMAPIPROP
			OpenNextSelectedItemProp(_Inout_opt_ int* iCurItem, __mfcmapiModifyEnum bModify) const;
			_Check_return_ std::vector<int> GetSelectedItemNums() const;
			_Check_return_ std::vector<sortlistdata::sortListData*> GetSelectedItemData() const;
			_Check_return_ sortlistdata::sortListData* GetFirstSelectedItemData() const;

			_Check_return_ HRESULT ApplyRestriction() const;
			_Check_return_ LPMAPIPROP DefaultOpenItemProp(int iItem, __mfcmapiModifyEnum bModify) const;
			void NotificationOff();
			void NotificationOn();
			void RefreshTable();
			void OnCancelTableLoad();
			void OnOutputTable(const std::wstring& szFileName) const;
			void SetSortTable(_In_ LPSSortOrderSet lpSortOrderSet, ULONG ulFlags) const;
			void SetUIColumns(_In_ LPSPropTagArray lpTags);
			_Check_return_ bool IsLoading() const;
			void ClearLoading();
			void SetRestriction(_In_opt_ const _SRestriction* lpRes);
			_Check_return_ const _SRestriction* GetRestriction() const;
			_Check_return_ __mfcmapiRestrictionTypeEnum GetRestrictionType() const;
			void SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType);
			_Check_return_ ULONG GetContainerType() const;
			_Check_return_ bool IsAdviseSet() const;
			_Check_return_ bool IsContentsTableSet() const;
			void DoSetColumns(bool bAddExtras, bool bDisplayEditor);
			void GetStatus();
			bool bAbortLoad() { return m_bAbortLoad != 0; }

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
			_Check_return_ int FindRow(_In_ LPSBinary lpInstance) const;
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

			ULONG m_ulDisplayFlags{dfNormal};
			LONG volatile m_bAbortLoad{false};
			std::thread m_LoadThreadHandle;
			dialog::CContentsTableDlg* m_lpHostDlg{nullptr};
			cache::CMapiObjects* m_lpMapiObjects{nullptr};
			std::vector<columns::TagNames> m_lpDefaultDisplayColumns;
			LPSPropTagArray m_sptDefaultDisplayColumnTags{nullptr};
			ULONG m_ulHeaderColumns{0};
			ULONG_PTR m_ulAdviseConnection{0};
			ULONG m_ulDisplayNameColumn{NODISPLAYNAME};
			UINT m_nIDContextMenu{0};
			bool m_bIsAB{false};
			bool m_bInLoadOp{false};
			const _SRestriction* m_lpRes{nullptr};
			ULONG m_ulContainerType{NULL};
			mapi::adviseSink* m_lpAdviseSink{nullptr};
			LPMAPITABLE m_lpContentsTable{nullptr};

			__mfcmapiRestrictionTypeEnum m_RestrictionType{mfcmapiNO_RESTRICTION};

			DECLARE_MESSAGE_MAP()
		};
	} // namespace sortlistctrl
} // namespace controls