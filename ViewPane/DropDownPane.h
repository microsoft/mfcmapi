#pragma once
// DropDownPane.h : header file

#include "ViewPane.h"

ViewPane* CreateDropDownPane(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, bool bReadOnly);
ViewPane* CreateDropDownArrayPane(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) LPNAME_ARRAY_ENTRY lpnaeDropList, bool bReadOnly);
ViewPane* CreateDropDownGuidPane(UINT uidLabel, bool bReadOnly);

class DropDownPane : public ViewPane
{
public:
	DropDownPane(UINT uidLabel, bool bReadOnly, ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, _In_opt_count_(ulDropList) LPNAME_ARRAY_ENTRY lpnaeDropList, bool bGUID);

	virtual bool IsType(__ViewTypes vType);
	virtual void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc);
	virtual void SetWindowPos(int x, int y, int width, int height);
	virtual void CommitUIValues();
	virtual ULONG GetFlags();
	virtual int GetMinWidth(_In_ HDC hdc);
	virtual int GetFixedHeight();
	virtual int GetLines();

	void SetDropDownSelection(_In_ wstring szText);

	void InsertDropString(int iRow, _In_ wstring szText);
	_Check_return_ CString GetDropStringUseControl();
	_Check_return_ int GetDropDownSelection();
	_Check_return_ DWORD_PTR GetDropDownSelectionValue();
	_Check_return_ bool GetSelectedGUID(bool bByteSwapped, _In_ LPGUID lpSelectedGUID);
	_Check_return_ int DropDownPane::GetDropDown();
	_Check_return_ DWORD_PTR DropDownPane::GetDropDownValue();

protected:
	CComboBox m_DropDown;
	DWORD_PTR m_iDropSelectionValue;

	void SetSelection(DWORD_PTR iSelection);

private:
	UINT* m_lpuidDropList;
	ULONG m_ulDropList; // count of entries in szDropList
	LPNAME_ARRAY_ENTRY m_lpnaeDropList;
	int m_iDropSelection;
	CString m_lpszSelectionString;
	bool m_bGUID;
};