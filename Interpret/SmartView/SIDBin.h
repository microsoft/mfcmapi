#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	class SIDBin : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockBytes m_SIDbin;
	};
} // namespace smartview