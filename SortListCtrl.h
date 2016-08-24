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

void FreeSortListData(_In_ SortListData* lpData);

class CSortListCtrl : public CListCtrl
{
public:
	CSortListCtrl();
	virtual ~CSortListCtrl();

	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();

	// Exported manipulation functions
	_Check_return_ HRESULT       Create(_In_ CWnd* pCreateParent, ULONG ulFlags, UINT nID, bool bImages);
	void          AutoSizeColumns(bool bMinWidth);
	void          DeleteAllColumns(bool bShutdown = false);
	void          SetSelectedItem(int iItem);
	void          SortClickedColumn();
	_Check_return_ SortListData* InsertRow(int iRow, wstring szText);
	void SetItemText(int nItem, int nSubItem, _In_z_ LPCTSTR lpszText);
	void SetItemTextA(int nItem, int nSubItem, _In_z_ LPCSTR lpszText);
	void SetItemTextW(int nItem, int nSubItem, _In_z_ LPCWSTR lpszText);
	wstring GetItemText(_In_ int nItem, _In_ int nSubItem) const;
	void AllowEscapeClose();

protected:
	void          MySetRedraw(bool bRedraw);
	_Check_return_ SortListData* InsertRow(int iRow, wstring szText, int iIndent, int iImage);
	void          FakeClickColumn(int iColumn, bool bSortUp);

	// protected since derived classes need to call the base implementation
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

private:
	// Overrides from base class
	_Check_return_ UINT OnGetDlgCode();

	void OnColumnClick(int iColumn);
	void OnDeleteAllItems(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void OnDeleteItem(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void AutoSizeColumn(int iColumn, int iMaxWidth, int iMinWidth);
	void OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	_Check_return_ static int CALLBACK MyCompareProc(_In_ LPARAM lParam1, _In_ LPARAM lParam2, _In_ LPARAM lParamSort);

	LONG		m_cRef;
	int			m_iRedrawCount;
	CImageList	m_ImageList;
	bool		m_bHaveSorted;
	bool		m_bHeaderSubclassed;
	CSortHeader	m_cSortHeader;
	int			m_iClickedColumn;
	bool		m_bSortUp;
	int			m_iItemCurHover;
	bool		m_bAllowEscapeClose;

	DECLARE_MESSAGE_MAP()
};