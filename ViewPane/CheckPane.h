#pragma once
#include "ViewPane.h"

class CheckPane : public ViewPane
{
public:
	CheckPane() : m_bCheckValue(false), m_bCommitted(false) {}

	static CheckPane* Create(UINT uidLabel, bool bVal, bool bReadOnly);
	bool GetCheck() const;

private:
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	void CommitUIValues() override;
	int GetMinWidth(_In_ HDC hdc) override;
	int GetFixedHeight() override;

	CButton m_Check;
	bool m_bCheckValue;
	bool m_bCommitted;
};