#include "stdafx.h"
#include "CollapsibleTextPane.h"

static wstring CLASS = L"CollapsibleTextPane";

CollapsibleTextPane* CollapsibleTextPane::Create(UINT uidLabel, bool bReadOnly)
{
	auto pane = new CollapsibleTextPane();
	if (pane)
	{
		pane->SetMultiline();
		pane->SetLabel(uidLabel, bReadOnly);
		pane->m_bCollapsible = true;
	}

	return pane;
}
