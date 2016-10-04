#pragma once
#include "ViewPane.h"
#include "SortList/SortListCtrl.h"

struct __ListButtons
{
	UINT uiButtonID;
};
#define NUMLISTBUTTONS 7

class CEditor;

class ListPane : public ViewPane
{
public:
	static ViewPane* Create(UINT uidLabel, bool bAllowSort, bool bReadOnly, LPVOID lpEdit);

	ULONG HandleChange(UINT nID) override;
	void SetListString(ULONG iListRow, ULONG iListCol, wstring szListString);
	_Check_return_ SortListData* InsertRow(int iRow, wstring szText) const;
	void ClearList();
	void ResizeList(bool bSort);
	_Check_return_ ULONG GetItemCount() const;
	_Check_return_ SortListData* GetItemData(int iRow) const;
	_Check_return_ SortListData* GetSelectedListRowData() const;
	void InsertColumn(int nCol, UINT uidText);
	void SetColumnType(int nCol, ULONG ulPropType) const;
	void UpdateListButtons();
	_Check_return_ bool OnEditListEntry();
	wstring GetItemText(_In_ int nItem, _In_ int nSubItem) const;

private:
	ListPane(bool bAllowSort, CEditor* lpEdit);

	bool IsType(__ViewTypes vType) override;
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	void CommitUIValues() override;
	ULONG GetFlags() override;
	int GetMinWidth(_In_ HDC hdc) override;
	int GetFixedHeight() override;
	int GetLines() override;

	void SwapListItems(ULONG ulFirstItem, ULONG ulSecondItem);
	void OnMoveListEntryUp();
	void OnMoveListEntryDown();
	void OnMoveListEntryToTop();
	void OnMoveListEntryToBottom();
	void OnAddListEntry();
	void OnDeleteListEntry(bool bDoDirty);

	CSortListCtrl m_List;
	CButton m_ButtonArray[NUMLISTBUTTONS];

	bool m_bDirty;
	bool m_bAllowSort;
	int m_iButtonWidth;

	// TODO: Figure out a way to do this that doesn't involve caching the edit control
	CEditor* m_lpEdit;
};