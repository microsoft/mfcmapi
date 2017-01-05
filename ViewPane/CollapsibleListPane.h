#pragma once
#include "ListPane.h"

class CollapsibleListPane : public ListPane
{
public:
	static CollapsibleListPane* Create(UINT uidLabel, bool bAllowSort, bool bReadOnly, DoListEditCallback callback);
};