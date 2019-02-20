#pragma once
// Contents table output for MrMAPI

namespace cli {
	struct MYOPTIONS;
}

void DoContents(_In_ cli::MYOPTIONS ProgOpts, LPMDB lpMDB, LPMAPIFOLDER lpFolder);
void DoMSG(_In_ cli::MYOPTIONS ProgOpts);