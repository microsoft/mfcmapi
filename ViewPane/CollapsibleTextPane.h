#pragma once
// CollapsibleTextPane.h : header file

#include "ViewPane.h"
#include "TextPane.h"

ViewPane* CreateCollapsibleTextPane(UINT uidLabel, bool bReadOnly);

class CollapsibleTextPane : public TextPane
{
public:
	CollapsibleTextPane();

	bool IsType(__ViewTypes vType) override;
	ULONG GetFlags() override;
	void SetWindowPos(int x, int y, int width, int height) override;
	int GetFixedHeight() override;
	int GetLines() override;
};