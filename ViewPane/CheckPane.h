#pragma once
// CheckPane.h : header file

#include "ViewPane.h"

ViewPane* CreateCheckPane(UINT uidLabel, bool bVal, bool bReadOnly);

class CheckPane : public ViewPane
{
public:
	CheckPane(UINT uidLabel, bool bReadOnly, bool bCheck);

	virtual bool IsType(__ViewTypes vType);
	virtual void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc);
	virtual void SetWindowPos(int x, int y, int width, int height);
	virtual void CommitUIValues();
	virtual ULONG GetFlags();
	virtual int GetMinWidth(_In_ HDC hdc);
	virtual int GetFixedHeight();
	virtual int GetLines();

	bool GetCheck();
	bool GetCheckUseControl();

private:
	CButton m_Check;

	bool m_bCheckValue;
};