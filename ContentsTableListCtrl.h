#pragma once

class CMapiObjects;
class CAdviseSink;
class CContentsTableDlg;
class CBaseDialog;

#include "SortListCtrl.h"
#include "enums.h"

/////////////////////////////////////////////////////////////////////////////
// CContentsTableListCtrl window

class CContentsTableListCtrl : public CSortListCtrl
{
friend class CContentsTableDlg;
public:
	CContentsTableListCtrl(
		CWnd* pCreateParent,
		CMapiObjects *lpMapiObjects,
		LPSPropTagArray	sptExtraColumnTags,
		ULONG ulNumExtraDisplayColumns,
		TagNames *lpExtraDisplayColumns,
		CContentsTableDlg *lpHostDlg);
	virtual ~CContentsTableListCtrl();

	HRESULT			GetSelectedItemEIDs(LPENTRYLIST* lppEntryIDs);
	SortListData*	GetNextSelectedItemData(int *iCurItem);
	HRESULT			OpenNextSelectedItemProp(int *iCurItem,__mfcmapiModifyEnum bModify,LPMAPIPROP* lppProp);

	HRESULT	SetContentsTable(
		LPMAPITABLE lpContentsTable,
		__mfcmapiDeletedItemsEnum bShowingDeletedItems,
		ULONG ulContainerType);

	void	OnCancelTableLoad();
	HRESULT RefreshItem(int iRow,LPSRow lpsRowData, BOOL bItemExists);

	//For my thread
	HRESULT ApplyRestriction();
	ULONG	GetThrottle(ULONG ulTotalRowCount);
	HRESULT DoFindRows(LPSRowSet *pRows);
	HRESULT DoSetColumns(BOOL bAddExtras, BOOL bDisplayEditor);

	HRESULT SetUIColumns(LPSPropTagArray lpTags);

	BOOL	m_bInLoadOp;

	//properties used externally - should have accessors
	LPSRestriction					m_lpRes;
	__mfcmapiRestrictionTypeEnum	m_RestrictionType;

	LPMAPITABLE			m_lpContentsTable;
protected:
	// Generated message map functions
	//{{AFX_MSG(CContentsTableListCtrl)
	afx_msg void OnItemChanged(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg LRESULT		msgOnAddItem(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnModifyItem(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnRefreshTable(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnThreadAddItem(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnSetAbort(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnGetAbort(WPARAM wParam, LPARAM lParam);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

private:
	HRESULT AddColumn(UINT uidHeaderName, ULONG ulCurHeaderCol, ULONG ulCurTagArrayRow, ULONG ulPropTag);
	HRESULT AddColumns(LPSPropTagArray lpCurColTagArray);
	HRESULT	AddItemToListBox(int iRow, LPSRow lpsRowToAdd);
	void	BuildDataItem(LPSRow lpsRowData,SortListData* lpData);
	HRESULT DefaultOpenItemProp(int iItem,__mfcmapiModifyEnum bModify,LPMAPIPROP* lppProp);
	int		FindRow(LPSBinary lpInstance);
	int		GetNextSelectedItemNum(int *iCurItem);
	HRESULT	LoadContentsTableIntoView();
	void	NotificationOff();
	HRESULT	NotificationOn();
	void	OnOutputTable(LPCTSTR szFileName);
	HRESULT	RefreshTable();
	void	SelectAll();
	void	SetRowStrings(int iRow,LPSRow lpsRowData);
	HRESULT SetSortTable(LPSSortOrderSet lpSortOrderSet, ULONG ulFlags);
	HRESULT DoExpandCollapse();


	__mfcmapiDeletedItemsEnum	m_bShowingDeletedItems;

	BOOL				m_bAbortLoad;

	HANDLE				m_LoadThreadHandle;

	CContentsTableDlg*	m_lpHostDlg;
	CMapiObjects*		m_lpMapiObjects;
	CAdviseSink*		m_lpAdviseSink;
	TagNames*			m_lpExtraDisplayColumns;
	LPSPropTagArray		m_sptExtraColumnTags;

	ULONG				m_ulHeaderColumns;
	ULONG_PTR			m_ulAdviseConnection;
	ULONG				m_ulNumExtraDisplayColumns;
	ULONG				m_ulDisplayNameColumn;
	ULONG				m_ulContainerType;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
