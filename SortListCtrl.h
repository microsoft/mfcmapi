#pragma once
// CSortListCtrl window

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
	SORTLIST_BINARY,
	SORTLIST_TREENODE
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

void FreeSortListData(SortListData* lpData);

class CSortListCtrl : public CListCtrl
{
public:
	CSortListCtrl();
	virtual ~CSortListCtrl();

	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();

	// Exported manipulation functions
	HRESULT       Create(CWnd* pCreateParent, ULONG ulFlags, UINT nID, BOOL bImages);
	void          AutoSizeColumns();
	void          DeleteAllColumns(BOOL bShutdown = false);
	void          SetSelectedItem(int iItem);
	void          SortClickedColumn();
	SortListData* InsertRow(int iRow, LPTSTR szText);
	BOOL          SetItemText(int nItem, int nSubItem, LPCTSTR lpszText);

protected:
	void          MySetRedraw(BOOL bRedraw);
	SortListData* InsertRow(int iRow, LPTSTR szText, int iIndent, int iImage);
	void          FakeClickColumn(int iColumn, BOOL bSortUp);

	// protected since derived classes need to call the base implementation
	virtual LRESULT	WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

private:
	// Overrides from base class
	UINT OnGetDlgCode();

	void OnColumnClick(int iColumn);
	void OnDeleteAllItems(NMHDR* pNMHDR, LRESULT* pResult);
	void OnDeleteItem(NMHDR* pNMHDR, LRESULT* pResult);
	void AutoSizeColumn(int iColumn, int iMinWidth, int iMaxWidth);
	static int CALLBACK MyCompareProc(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort);

	LONG		m_cRef;
	int			m_iRedrawCount;
	CImageList	m_ImageList;
	BOOL		m_bHaveSorted;
	BOOL		m_bHeaderSubclassed;
	CSortHeader	m_cSortHeader;
	int			m_iClickedColumn;
	BOOL		m_bSortUp;

	DECLARE_MESSAGE_MAP()
};