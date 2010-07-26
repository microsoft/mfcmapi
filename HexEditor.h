#pragma once
// HexEditor.h : header file

#include "Editor.h"
#include "ParentWnd.h"

class CHexEditor : public CEditor
{
public:
	CHexEditor(
		_In_ CParentWnd* pParentWnd);
	virtual ~CHexEditor();

private:
	_Check_return_ ULONG HandleChange(UINT nID);
	void UpdateParser();

	void OnOK();
	void OnCancel();
	_Check_return_ virtual BOOL CheckAutoCenter();
	CParentWnd* m_lpNonModalParent;
	CWnd* m_hwndCenteringWindow;
};