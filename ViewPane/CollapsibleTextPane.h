#pragma once
// CollapsibleTextPane.h : header file

#include "ViewPane.h"
#include "TextPane.h"

class CollapsibleTextPane : public TextPane
{
public:
	static ViewPane* Create(UINT uidLabel, bool bReadOnly);

private:
	CollapsibleTextPane();

	bool IsType(__ViewTypes vType) override;
	ULONG GetFlags() override;
	void SetWindowPos(int x, int y, int width, int height) override;
	int GetFixedHeight() override;
	int GetLines() override;
};