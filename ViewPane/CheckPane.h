#pragma once
// CheckPane.h : header file

#include "ViewPane.h"

class CheckPane : public ViewPane
{
public:
	static CheckPane* Create(UINT uidLabel, bool bVal, bool bReadOnly);
	bool GetCheck() const;
	bool GetCheckUseControl() const;

private:
	bool IsType(__ViewTypes vType) override;
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	void CommitUIValues() override;
	ULONG GetFlags() override;
	int GetMinWidth(_In_ HDC hdc) override;
	int GetFixedHeight() override;
	int GetLines() override;

	CButton m_Check;
	bool m_bCheckValue;
};