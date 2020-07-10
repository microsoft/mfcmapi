#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	constexpr WORD kwBaseOffset = 0xAC00; // Hangul char range (AC00-D7AF)

	class encodeEntryID : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockBytes> entryID = emptyBB();
	};
} // namespace smartview