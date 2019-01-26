#pragma once
// Acl table dumping for MrMAPI

namespace cli {
	struct MYOPTIONS;
}

void DoAcls(_In_ cli::MYOPTIONS ProgOpts);
void DumpExchangeTable(_In_ ULONG ulPropTag, _In_ LPMAPIFOLDER lpFolder);