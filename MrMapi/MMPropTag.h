#pragma once
// Prop tag parsing for MrMAPI

namespace cli
{
	struct MYOPTIONS;
}

void DoPropTags(_In_ const cli::MYOPTIONS& ProgOpts);
void DoGUIDs(_In_ const cli::MYOPTIONS& ProgOpts);
void DoFlagSearch(_In_ const cli::MYOPTIONS& ProgOpts);