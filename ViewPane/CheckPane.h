#pragma once
// CheckPane.h : header file

#include "ViewPane.h"

ViewPane* CreateCheckPane(UINT uidLabel, bool bVal, bool bReadOnly);

class CheckPane : public ViewPane
{
public:
	CheckPane(UINT uidLabel, bool bReadOnly, bool bCheck);

	bool IsType(__ViewTypes vType) override;
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	void CommitUIValues() override;
	ULONG GetFlags() override;
	int GetMinWidth(_In_ HDC hdc) override;
	int GetFixedHeight() override;
	int GetLines() override;

	bool GetCheck() const;
	bool GetCheckUseControl() const;

private:
	CButton m_Check;

	bool m_bCheckValue;
};