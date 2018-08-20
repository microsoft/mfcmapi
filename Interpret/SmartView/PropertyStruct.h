#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <MAPIDefs.h>

namespace smartview
{
	class PropertyStruct : public SmartViewParser
	{
	public:
		PropertyStruct();

	private:
		void Parse() override;
		void ParseBlocks() override;

		std::vector<SPropValueStruct> m_Prop;
	};

	_Check_return_ block SPropValueStructToBlock(std::vector<SPropValueStruct> props);
	_Check_return_ std::wstring PropsToString(DWORD PropCount, LPSPropValue Prop);
}