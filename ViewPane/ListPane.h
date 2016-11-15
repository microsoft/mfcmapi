#pragma once
#include "ViewPane.h"
#include "Controls/SortList/SortListCtrl.h"
#include <functional>

#define NUMLISTBUTTONS 7

typedef std::function<bool(ULONG, int, SortListData*)> DoListEditCallback;

class ListPane : public ViewPane
{
public:
	static ListPane* Create(UINT uidLabel, bool bAllowSort, bool bReadOnly, DoListEditCallback callback);

	ULONG HandleChange(UINT nID) override;
	void SetListString(ULONG iListRow, ULONG iListCol, const wstring& szListString);
	_Check_return_ SortListData* InsertRow(int iRow, const wstring& szText) const;
	void ClearList();
	void ResizeList(bool bSort);
	_Check_return_ ULONG GetItemCount() const;
	_Check_return_ SortListData* GetItemData(int iRow) const;
	_Check_return_ SortListData* GetSelectedListRowData() const;
	void InsertColumn(int nCol, UINT uidText);
	void SetColumnType(int nCol, ULONG ulPropType) const;
	_Check_return_ bool OnEditListEntry();
	wstring GetItemText(_In_ int nItem, _In_ int nSubItem) const;

private:
	ListPane();
	void Setup(bool bAllowSort, DoListEditCallback callback);
	void UpdateButtons() override;

	ULONG GetFlags() override;
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	void CommitUIValues() override;
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

	DoListEditCallback m_callback;
};