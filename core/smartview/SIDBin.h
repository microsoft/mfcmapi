#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	class SIDBin : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockBytes> m_SIDbin = emptyBB();
	};
} // namespace smartview