#pragma once

#include <Dialogs/Editors/Editor.h>
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
	_Check_return_ ULONG HandleChange(UINT nID) override;
	void OnEditAction1() override;
	void OnEditAction2() override;
	void OnEditAction3() override;
	void UpdateParser() const;

	void OnOK() override;
	void OnCancel() override;

	CMapiObjects* m_lpMapiObjects;
};