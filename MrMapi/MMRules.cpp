#include <StdAfx.h>
#include <MrMapi/MMRules.h>
#include <MrMapi/mmcli.h>
#include <MrMapi/MMAcls.h>

void DoRules(_In_ cli::MYOPTIONS ProgOpts) { DumpExchangeTable(PR_RULES_TABLE, ProgOpts.lpFolder); }
