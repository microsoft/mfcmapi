#pragma once
// DropDownPane.h : header file

#include "ViewPane.h"

class DropDownPane : public ViewPane
{
public:
	static ViewPane* Create(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, bool bReadOnly);
	static ViewPane* CreateArray(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) LPNAME_ARRAY_ENTRY lpnaeDropList, bool bReadOnly);
	static ViewPane* CreateGuid(UINT uidLabel, bool bReadOnly);

	DropDownPane(ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, _In_opt_count_(ulDropList) LPNAME_ARRAY_ENTRY lpnaeDropList, bool bGUID);

	void SetDropDownSelection(_In_ wstring szText);
	void InsertDropString(int iRow, _In_ wstring szText, ULONG ulValue);
	_Check_return_ wstring GetDropStringUseControl() const;
	_Check_return_ int GetDropDownSelection() const;
	_Check_return_ DWORD_PTR GetDropDownSelectionValue() const;
	_Check_return_ bool GetSelectedGUID(bool bByteSwapped, _In_ LPGUID lpSelectedGUID) const;
	_Check_return_ int GetDropDown() const;
	_Check_return_ DWORD_PTR GetDropDownValue() const;

protected:
	bool IsType(__ViewTypes vType) override;
	ULONG GetFlags() override;
	void SetSelection(DWORD_PTR iSelection);
	void DoInit(int iControl, _In_ CWnd* pParent, _In_ HDC hdc);

	CComboBox m_DropDown;
	DWORD_PTR m_iDropSelectionValue;


private:
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	void CommitUIValues() override;
	int GetMinWidth(_In_ HDC hdc) override;
	int GetFixedHeight() override;
	int GetLines() override;

	ULONG m_ulDropList; // count of entries in m_lpuidDropList or m_lpnaeDropList
	UINT* m_lpuidDropList;
	LPNAME_ARRAY_ENTRY m_lpnaeDropList;
	int m_iDropSelection;
	wstring m_lpszSelectionString;
	bool m_bGUID;
};