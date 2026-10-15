#pragma once
#include <UI/ViewPane/ViewPane.h>

namespace viewpane
{
	class SectionPane final : public ViewPane
	{
	public:
		SectionPane(int paneID, UINT uidLabel);

	private:
		void CommitUIValues() override {}
		int GetFixedHeight() override;
	};
} // namespace viewpane
