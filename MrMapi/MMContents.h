#pragma once
// Contents table output for MrMAPI

namespace cli
{
	struct OPTIONS;
}

void DoContents(_In_ cli::OPTIONS ProgOpts, LPMDB lpMDB, LPMAPIFOLDER lpFolder);
void DoMSG(_In_ cli::OPTIONS ProgOpts);
