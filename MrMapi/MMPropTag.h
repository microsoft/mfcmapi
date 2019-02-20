#pragma once
// Prop tag parsing for MrMAPI

namespace cli
{
	struct OPTIONS;
}

void DoPropTags(_In_ const cli::OPTIONS& ProgOpts);
void DoGUIDs();
void DoFlagSearch();