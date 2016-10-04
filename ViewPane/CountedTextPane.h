#pragma once
// CountedTextPane.h : header file

#include "ViewPane.h"
#include "TextPane.h"

class CountedTextPane : public TextPane
{
public:
	static CountedTextPane* CountedTextPane::Create(UINT uidLabel, bool bReadOnly, UINT uidCountLabel);
	void SetCount(size_t iCount);

private:
	bool IsType(__ViewTypes vType) override;
	ULONG GetFlags() override;
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	int GetMinWidth(_In_ HDC hdc) override;
	int GetFixedHeight() override;
	int GetLines() override;

	CEdit m_Count; // The display of the count
	UINT m_uidCountLabel; // UID for the name of the count
	wstring m_szCountLabel; // String name of the count
	int m_iCountLabelWidth; // The width of the string
	size_t m_iCount; // The numeric count
};