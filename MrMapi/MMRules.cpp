#include <StdAfx.h>
#include <MrMapi/MrMAPI.h>
#include <MrMapi/MMRules.h>
#include <MrMapi/MMAcls.h>

void DoRules(_In_ MYOPTIONS ProgOpts) { DumpExchangeTable(PR_RULES_TABLE, ProgOpts.lpFolder); }