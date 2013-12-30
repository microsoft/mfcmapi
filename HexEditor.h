#pragma once
// HexEditor.h : header file

#include "Editor.h"
#include "ParentWnd.h"
#include "MapiObjects.h"

class CHexEditor : public CEditor
{
public:
	CHexEditor(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects);
	virtual ~CHexEditor();

private:
	_Check_return_ ULONG HandleChange(UINT nID);
	void OnEditAction1();
	void OnEditAction2();
	void OnEditAction3();
	void UpdateParser();

	void OnOK();
	void OnCancel();

	CMapiObjects* m_lpMapiObjects;
};