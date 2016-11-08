#pragma once
#include "ViewPane.h"
#include "TextPane.h"

class CollapsibleTextPane : public TextPane
{
public:
	static CollapsibleTextPane* Create(UINT uidLabel, bool bReadOnly);

private:
	bool IsType(__ViewTypes vType) override;
	ULONG GetFlags() override;
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	int GetFixedHeight() override;
	int GetLines() override;
};