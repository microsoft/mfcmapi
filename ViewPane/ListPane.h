#pragma once
// ListPane.h : header file

#include "ViewPane.h"
#include "..\SortListCtrl.h"

ViewPane* CreateListPane(UINT uidLabel, bool bAllowSort, bool bReadOnly, LPVOID lpEdit);

struct __ListButtons
{
	UINT uiButtonID;
};
#define NUMLISTBUTTONS 7

class CEditor;

class ListPane : public ViewPane
{
public:
	ListPane(UINT uidLabel, bool bReadOnly, bool bAllowSort, CEditor* lpEdit);

	virtual bool IsType(__ViewTypes vType);
	virtual void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc);
	virtual void SetWindowPos(int x, int y, int width, int height);
	virtual ULONG GetFlags();
	virtual int GetMinWidth(_In_ HDC hdc);
	virtual int GetFixedHeight();
	virtual int GetLines();
	virtual ULONG HandleChange(UINT nID);

	void SetListString(ULONG iListRow, ULONG iListCol, wstring szListString);

	_Check_return_ SortListData* InsertRow(int iRow, _In_z_ LPCTSTR szText);
	void ClearList();
	void ResizeList(bool bSort);
	_Check_return_ ULONG GetItemCount();
	_Check_return_ SortListData* GetItemData(int iRow);
	_Check_return_ SortListData* GetSelectedListRowData();
	void InsertColumn(int nCol, UINT uidText);
	void SetColumnType(int nCol, ULONG ulPropType);
	void UpdateListButtons();
	void SwapListItems(ULONG ulFirstItem, ULONG ulSecondItem);
	void OnMoveListEntryUp();
	void OnMoveListEntryDown();
	void OnMoveListEntryToTop();
	void OnMoveListEntryToBottom();
	void OnAddListEntry();
	void OnDeleteListEntry(bool bDoDirty);
	_Check_return_ bool OnEditListEntry();
	CString GetItemText(_In_ int nItem, _In_ int nSubItem);

private:
	CSortListCtrl m_List;
	CButton m_ButtonArray[NUMLISTBUTTONS];

	bool m_bDirty;
	bool m_bAllowSort;
	int m_iButtonWidth;

	// TODO: Figure out a way to do this that doesn't involve caching the edit control
	CEditor* m_lpEdit;
};