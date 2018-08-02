#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	class GlobalObjectId : public SmartViewParser
	{
	private:
		void Parse() override;
	};
}