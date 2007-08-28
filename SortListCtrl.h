#pragma once

#include "enums.h"

#include "SortHeader.h"

enum __SortListDataTypes
{
	SORTLIST_UNKNOWN = 0,
	SORTLIST_CONTENTS,
	SORTLIST_PROP,
	SORTLIST_MVPROP,
	SORTLIST_TAGARRAY,
	SORTLIST_RES,
	SORTLIST_COMMENT,
	SORTLIST_BINARY
};

// This enum maps to icons.bmp
// Reorder at your own risk!!!
enum __SortListIconNames
{
	slIconDefault = 0,
	slIconNodeCollapsed,
	slIconNodeExpanded,
	slIconUNSPECIFIED,
	slIconNULL,
	slIconI2,
	slIconLONG,
	slIconR4,
	slIconDOUBLE,
	slIconCURRENCY,
	slIconAPPTIME,
	slIconERROR,
	slIconBOOLEAN,
	slIconOBJECT,
	slIconI8,
	slIconSTRING8,
	slIconUNICODE,
	slIconSYSTIME,
	slIconCLSID,
	slIconBINARY,
	slIconMV_I2,
	slIconMV_LONG,
	slIconMV_R4,
	slIconMV_DOUBLE,
	slIconMV_CURRENCY,
	slIconMV_APPTIME,
	slIconMV_SYSTIME,
	slIconMV_STRING8,
	slIconMV_BINARY,
	slIconMV_UNICODE,
	slIconMV_CLSID,
	slIconMV_I8,
	slIconSRESTRICTION,
	slIconACTIONS,
	slIconMAPI_STORE,
	slIconMAPI_ADDRBOOK,
	slIconMAPI_FOLDER,
	slIconMAPI_ABCONT,
	slIconMAPI_MESSAGE,
	slIconMAPI_MAILUSER,
	slIconMAPI_ATTACH,
	slIconMAPI_DISTLIST,
	slIconMAPI_PROFSECT,
	slIconMAPI_STATUS,
	slIconMAPI_SESSION,
	slIconMAPI_FORMINFO,
};

/////////////////////////////////////////////////////////////////////////////
// CContentsTableListCtrl window

class CSortListCtrl : public CListCtrl
{
public:
	CSortListCtrl();
	virtual ~CSortListCtrl();

	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();

	HRESULT Create(CWnd* pCreateParent, ULONG ulFlags, UINT nID, BOOL bImages);

	void	MySetRedraw(BOOL bRedraw);
	int		m_iClickedColumn;

	void	AutoSizeColumns();
	void	DeleteAllColumns();
	void	SetSelectedItem(int iItem);
	void	SortColumn(ULONG iColumn);
	SortListData* InsertRow(int iRow, LPTSTR szText);
	SortListData* InsertRow(int iRow, LPTSTR szText, int iIndent, int iImage);
	void	GetSelectedItems(int* cSelected, SortListData*** lpSelected);

protected:
	// Generated message map functions
	//{{AFX_MSG(CContentsTableListCtrl)
	afx_msg void OnColumnClick(int iColumn);
	afx_msg UINT OnGetDlgCode();
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnDeleteAllItems(NMHDR* pNMHDR, LRESULT* pResult);
	virtual afx_msg void OnDeleteItem(NMHDR* pNMHDR, LRESULT* pResult);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

	BOOL	m_bSortUp;
private:
	LONG	m_cRef;
	int		m_iRedrawCount;
	CImageList m_ImageList;

	BOOL	m_bHaveSorted;

	BOOL	m_bHeaderSubclassed;
	CSortHeader	m_cSortHeader;
	void	AutoSizeColumn(int iColumn, int iMinWidth, int iMaxWidth);

	static int CALLBACK MyCompareProc(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.