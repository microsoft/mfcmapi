#pragma once
// CollapsibleTextPane.h : header file

#include "ViewPane.h"
#include "TextPane.h"

ViewPane* CreateCollapsibleTextPane(UINT uidLabel, bool bReadOnly);

class CollapsibleTextPane : public TextPane
{
public:
	CollapsibleTextPane(UINT uidLabel, bool bReadOnly);

	virtual bool IsType(__ViewTypes vType);
	virtual ULONG GetFlags();
	virtual void SetWindowPos(int x, int y, int width, int height);
	virtual int GetFixedHeight();
	virtual int GetLines();
};