#pragma once
// HexEditor.h : header file

#include "Editor.h"

class CHexEditor : public CEditor
{
public:
	CHexEditor(
		CParentWnd* pParentWnd);
	virtual ~CHexEditor();

private:
	ULONG HandleChange(UINT nID);
	void UpdateParser();

	void OnOK();
	void OnCancel();
	CParentWnd* m_lpNonModalParent;
};