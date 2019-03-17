#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	class SIDBin : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::shared_ptr<blockBytes> m_SIDbin = emptyBB();
	};
} // namespace smartview