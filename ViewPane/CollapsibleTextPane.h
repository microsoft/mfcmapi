#pragma once
#include "TextPane.h"

class CollapsibleTextPane : public TextPane
{
public:
	static CollapsibleTextPane* Create(UINT uidLabel, bool bReadOnly);

private:
	void SetWindowPos(int x, int y, int width, int height) override;
	int GetFixedHeight() override;
};