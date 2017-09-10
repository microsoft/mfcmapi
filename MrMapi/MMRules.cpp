#include "stdafx.h"
#include "MrMAPI.h"
#include "MMRules.h"
#include "MMAcls.h"

void DoRules(_In_ MYOPTIONS ProgOpts)
{
	DumpExchangeTable(PR_RULES_TABLE, ProgOpts.lpFolder);
}