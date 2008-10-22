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
		CWnd* pCreateParent,
		CMapiObjects* lpMapiObjects,
		LPSPropTagArray	sptExtraColumnTags,
		ULONG ulNumExtraDisplayColumns,
		TagNames* lpExtraDisplayColumns,
		UINT nIDContextMenu,
		BOOL bIsAB,
		CContentsTableDlg* lpHostDlg);
	virtual ~CContentsTableListCtrl();

	// Initialization
	HRESULT	SetContentsTable(
		LPMAPITABLE lpContentsTable,
		ULONG ulDisplayFlags,
		ULONG ulContainerType);

	// Selected item accessors
	HRESULT			GetSelectedItemEIDs(LPENTRYLIST* lppEntryIDs);
	SortListData*	GetNextSelectedItemData(int* iCurItem);
	HRESULT			OpenNextSelectedItemProp(int* iCurItem,__mfcmapiModifyEnum bModify,LPMAPIPROP* lppProp);

	HRESULT ApplyRestriction();
	HRESULT DefaultOpenItemProp(int iItem,__mfcmapiModifyEnum bModify,LPMAPIPROP* lppProp);
	void	NotificationOff();
	HRESULT	NotificationOn();
	HRESULT	RefreshTable();
	void	OnCancelTableLoad();
	void	OnOutputTable(LPCTSTR szFileName);
	HRESULT SetSortTable(LPSSortOrderSet lpSortOrderSet, ULONG ulFlags);
	HRESULT SetUIColumns(LPSPropTagArray lpTags);
	BOOL	IsLoading();
	void	ClearLoading();
	void	SetRestriction(LPSRestriction lpRes);
	LPSRestriction GetRestriction();
	__mfcmapiRestrictionTypeEnum GetRestrictionType();
	void	SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType);
	ULONG	GetContainerType();
	BOOL	IsAdviseSet();
	BOOL	IsContentsTableSet();
	void	DoSetColumns(BOOL bAddExtras, BOOL bDisplayEditor, BOOL bQueryFlags, BOOL bDoRefresh);
	void	GetStatus();

private:
	// Overrides from base class
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void	OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	void	OnItemChanged(NMHDR* pNMHDR, LRESULT* pResult);

	HRESULT AddColumn(UINT uidHeaderName, ULONG ulCurHeaderCol, ULONG ulCurTagArrayRow, ULONG ulPropTag);
	HRESULT AddColumns(LPSPropTagArray lpCurColTagArray);
	HRESULT	AddItemToListBox(int iRow, LPSRow lpsRowToAdd);
	void	BuildDataItem(LPSRow lpsRowData,SortListData* lpData);
	HRESULT DoExpandCollapse();
	int		FindRow(LPSBinary lpInstance);
	int		GetNextSelectedItemNum(int *iCurItem);
	HRESULT	LoadContentsTableIntoView();
	HRESULT RefreshItem(int iRow,LPSRow lpsRowData, BOOL bItemExists);
	void	SelectAll();
	void	SetRowStrings(int iRow,LPSRow lpsRowData);

	// Custom messages
	LRESULT	msgOnAddItem(WPARAM wParam, LPARAM lParam);
	LRESULT	msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
	LRESULT	msgOnModifyItem(WPARAM wParam, LPARAM lParam);
	LRESULT	msgOnRefreshTable(WPARAM wParam, LPARAM lParam);
	LRESULT	msgOnThreadAddItem(WPARAM wParam, LPARAM lParam);
	LRESULT	msgOnClearAbort(WPARAM wParam, LPARAM lParam);
	LRESULT	msgOnGetAbort(WPARAM wParam, LPARAM lParam);

	ULONG				m_ulDisplayFlags;
	BOOL				m_bAbortLoad;
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
	BOOL				m_bIsAB;
	BOOL				m_bInLoadOp;
	LPSRestriction		m_lpRes;
	ULONG				m_ulContainerType;
	CAdviseSink*		m_lpAdviseSink;
	LPMAPITABLE			m_lpContentsTable;

	__mfcmapiRestrictionTypeEnum m_RestrictionType;

	DECLARE_MESSAGE_MAP()
};