#include "stdafx.h"
#include "CollapsibleListPane.h"

static wstring CLASS = L"CollapsibleListPane";

CollapsibleListPane* CollapsibleListPane::Create(UINT uidLabel, bool bAllowSort, bool bReadOnly, DoListEditCallback callback)
{
	auto pane = new CollapsibleListPane();
	if (pane)
	{
		pane->Setup(bAllowSort, callback);
		pane->SetLabel(uidLabel, bReadOnly);
		pane->m_bCollapsible = true;
	}

	return pane;
}