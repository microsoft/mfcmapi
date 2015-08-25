#pragma once
// ContentsTableListCtrl.h : header file

class CMapiObjects;
class CContentsTableDlg;
#include "SortListCtrl.h"
#include "enums.h"

class CContentsTableListCtrl : public CSortListCtrl
{
public:
	CContentsTableListCtrl(
		_In_ CWnd* pCreateParent,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPSPropTagArray sptExtraColumnTags,
		ULONG ulNumExtraDisplayColumns,
		_In_count_(ulNumExtraDisplayColumns) TagNames* lpExtraDisplayColumns,
		UINT nIDContextMenu,
		bool bIsAB,
		_In_ CContentsTableDlg* lpHostDlg);
	virtual ~CContentsTableListCtrl();

	// Initialization
	_Check_return_ HRESULT SetContentsTable(
		_In_opt_ LPMAPITABLE lpContentsTable,
		ULONG ulDisplayFlags,
		ULONG ulContainerType);

	// Selected item accessors
	_Check_return_ HRESULT       GetSelectedItemEIDs(_Deref_out_opt_ LPENTRYLIST* lppEntryIDs);
	_Check_return_ SortListData* GetNextSelectedItemData(_Inout_opt_ int* iCurItem);
	_Check_return_ HRESULT       OpenNextSelectedItemProp(_Inout_opt_ int* iCurItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppProp);

	_Check_return_ HRESULT ApplyRestriction();
	_Check_return_ HRESULT DefaultOpenItemProp(int iItem,__mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppProp);
	void    NotificationOff();
	_Check_return_ HRESULT NotificationOn();
	_Check_return_ HRESULT RefreshTable();
	void    OnCancelTableLoad();
	void    OnOutputTable(wstring szFileName);
	_Check_return_ HRESULT SetSortTable(_In_ LPSSortOrderSet lpSortOrderSet, ULONG ulFlags);
	_Check_return_ HRESULT SetUIColumns(_In_ LPSPropTagArray lpTags);
	_Check_return_ bool    IsLoading();
	void    ClearLoading();
	void    SetRestriction(_In_opt_ LPSRestriction lpRes);
	_Check_return_ LPSRestriction GetRestriction();
	_Check_return_ __mfcmapiRestrictionTypeEnum GetRestrictionType();
	void    SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType);
	_Check_return_ ULONG   GetContainerType();
	_Check_return_ bool    IsAdviseSet();
	_Check_return_ bool    IsContentsTableSet();
	void    DoSetColumns(bool bAddExtras, bool bDisplayEditor);
	void    GetStatus();

private:
	// Overrides from base class
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void    OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	void    OnItemChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void    OnContextMenu(_In_ CWnd *pWnd, CPoint pos);

	_Check_return_ HRESULT AddColumn(UINT uidHeaderName, ULONG ulCurHeaderCol, ULONG ulCurTagArrayRow, ULONG ulPropTag);
	_Check_return_ HRESULT AddColumns(_In_ LPSPropTagArray lpCurColTagArray);
	_Check_return_ HRESULT AddItemToListBox(int iRow, _In_ LPSRow lpsRowToAdd);
	void    BuildDataItem(_In_ LPSRow lpsRowData, _Inout_ SortListData* lpData);
	_Check_return_ HRESULT DoExpandCollapse();
	_Check_return_ int     FindRow(_In_ LPSBinary lpInstance);
	_Check_return_ int     GetNextSelectedItemNum(_Inout_opt_ int *iCurItem);
	_Check_return_ HRESULT LoadContentsTableIntoView();
	_Check_return_ HRESULT RefreshItem(int iRow, _In_ LPSRow lpsRowData, bool bItemExists);
	void    SelectAll();
	void    SetRowStrings(int iRow, _In_ LPSRow lpsRowData);

	// Custom messages
	_Check_return_ LRESULT msgOnAddItem(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnModifyItem(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnRefreshTable(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnThreadAddItem(WPARAM wParam, LPARAM lParam);

	ULONG				m_ulDisplayFlags;
	LONG volatile		m_bAbortLoad;
	HANDLE				m_LoadThreadHandle;
	CContentsTableDlg*	m_lpHostDlg;
	CMapiObjects*		m_lpMapiObjects;
	TagNames*			m_lpExtraDisplayColumns;
	LPSPropTagArray		m_sptExtraColumnTags;
	ULONG				m_ulHeaderColumns;
	ULONG_PTR			m_ulAdviseConnection;
	ULONG				m_ulNumExtraDisplayColumns;
	ULONG				m_ulDisplayNameColumn;
	UINT				m_nIDContextMenu;
	bool				m_bIsAB;
	bool				m_bInLoadOp;
	LPSRestriction		m_lpRes;
	ULONG				m_ulContainerType;
	CAdviseSink*		m_lpAdviseSink;
	LPMAPITABLE			m_lpContentsTable;

	__mfcmapiRestrictionTypeEnum m_RestrictionType;

	DECLARE_MESSAGE_MAP()
};