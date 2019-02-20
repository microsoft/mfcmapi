#pragma once
// Contents table output for MrMAPI

namespace cli
{
	struct OPTIONS;
}

void DoContents(LPMDB lpMDB, LPMAPIFOLDER lpFolder);
void DoMSG(_In_ cli::OPTIONS ProgOpts);
