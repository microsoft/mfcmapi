#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	class SIDBin : public smartViewParser
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockBytes> m_SIDbin = emptyBB();
	};
} // namespace smartview