#pragma once
#include <UI/Controls/SortList/SortListCtrl.h>
#include <Enums.h>

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
	namespace mapiui
	{
		class CAdviseSink;
	}
}

namespace controls
{
	namespace sortlistctrl
	{
		class CContentsTableListCtrl : public CSortListCtrl
		{
		public:
			CContentsTableListCtrl(
				_In_ CWnd* pCreateParent,
				_In_ cache::CMapiObjects* lpMapiObjects,
				_In_ LPSPropTagArray sptExtraColumnTags,
				_In_ const std::vector<TagNames>& lpExtraDisplayColumns,
				UINT nIDContextMenu,
				bool bIsAB,
				_In_ dialog::CContentsTableDlg* lpHostDlg);
			virtual ~CContentsTableListCtrl();

			// Initialization
			_Check_return_ HRESULT SetContentsTable(
				_In_opt_ LPMAPITABLE lpContentsTable,
				ULONG ulDisplayFlags,
				ULONG ulContainerType);

			// Selected item accessors
			_Check_return_ HRESULT GetSelectedItemEIDs(_Deref_out_opt_ LPENTRYLIST* lppEntryIDs) const;
			_Check_return_ sortlistdata::SortListData* GetSortListData(int iItem) const;
			_Check_return_ HRESULT OpenNextSelectedItemProp(_Inout_opt_ int* iCurItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppProp) const;
			_Check_return_ std::vector<int> GetSelectedItemNums() const;
			_Check_return_ std::vector<sortlistdata::SortListData*> GetSelectedItemData() const;
			_Check_return_ sortlistdata::SortListData* GetFirstSelectedItemData() const;

			_Check_return_ HRESULT ApplyRestriction() const;
			_Check_return_ HRESULT DefaultOpenItemProp(int iItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppProp) const;
			void NotificationOff();
			_Check_return_ HRESULT NotificationOn();
			_Check_return_ HRESULT RefreshTable();
			void OnCancelTableLoad();
			void OnOutputTable(const std::wstring& szFileName) const;
			_Check_return_ HRESULT SetSortTable(_In_ LPSSortOrderSet lpSortOrderSet, ULONG ulFlags) const;
			_Check_return_ HRESULT SetUIColumns(_In_ LPSPropTagArray lpTags);
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

		private:
			// Overrides from base class
			LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
			void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
			void OnItemChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
			void OnContextMenu(_In_ CWnd *pWnd, CPoint pos);

			_Check_return_ HRESULT AddColumn(UINT uidHeaderName, ULONG ulCurHeaderCol, ULONG ulCurTagArrayRow, ULONG ulPropTag);
			_Check_return_ HRESULT AddColumns(_In_ LPSPropTagArray lpCurColTagArray);
			_Check_return_ HRESULT AddItemToListBox(int iRow, _In_ LPSRow lpsRowToAdd);
			_Check_return_ HRESULT DoExpandCollapse();
			_Check_return_ int FindRow(_In_ LPSBinary lpInstance) const;
			_Check_return_ int GetNextSelectedItemNum(_Inout_opt_ int *iCurItem) const;
			_Check_return_ HRESULT LoadContentsTableIntoView();
			_Check_return_ HRESULT RefreshItem(int iRow, _In_ LPSRow lpsRowData, bool bItemExists);
			void SelectAll();
			void SetRowStrings(int iRow, _In_ LPSRow lpsRowData);

			// Custom messages
			_Check_return_ LRESULT msgOnAddItem(WPARAM wParam, LPARAM lParam);
			_Check_return_ LRESULT msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
			_Check_return_ LRESULT msgOnModifyItem(WPARAM wParam, LPARAM lParam);
			_Check_return_ LRESULT msgOnRefreshTable(WPARAM wParam, LPARAM lParam);
			_Check_return_ LRESULT msgOnThreadAddItem(WPARAM wParam, LPARAM lParam);

			ULONG m_ulDisplayFlags;
			LONG volatile m_bAbortLoad;
			HANDLE m_LoadThreadHandle;
			dialog::CContentsTableDlg* m_lpHostDlg;
			cache::CMapiObjects* m_lpMapiObjects;
			std::vector<TagNames> m_lpExtraDisplayColumns;
			LPSPropTagArray m_sptExtraColumnTags;
			ULONG m_ulHeaderColumns;
			ULONG_PTR m_ulAdviseConnection;
			ULONG m_ulDisplayNameColumn;
			UINT m_nIDContextMenu;
			bool m_bIsAB;
			bool m_bInLoadOp;
			const _SRestriction* m_lpRes;
			ULONG m_ulContainerType;
			mapi::mapiui::CAdviseSink* m_lpAdviseSink;
			LPMAPITABLE m_lpContentsTable;

			__mfcmapiRestrictionTypeEnum m_RestrictionType;

			DECLARE_MESSAGE_MAP()
		};
	}
}