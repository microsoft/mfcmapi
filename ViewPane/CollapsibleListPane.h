#pragma once
#include "ListPane.h"

class CollapsibleListPane : public ListPane
{
public:
	static CollapsibleListPane* Create(UINT uidLabel, bool bAllowSort, bool bReadOnly, DoListEditCallback callback);

private:
	ULONG GetFlags() override;
	void SetWindowPos(int x, int y, int width, int height) override;
	int GetFixedHeight() override;
	int GetLines() override;
};