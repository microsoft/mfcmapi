#pragma once
#include "TextPane.h"

class CollapsibleTextPane : public TextPane
{
public:
	static CollapsibleTextPane* Create(UINT uidLabel, bool bReadOnly);
};