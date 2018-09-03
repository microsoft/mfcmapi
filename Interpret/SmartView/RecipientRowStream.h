#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/PropertiesStruct.h>

namespace smartview
{
	struct ADRENTRYStruct
	{
		blockT<DWORD> ulReserved1;
		blockT<DWORD> cValues;
		PropertiesStruct rgPropVals;
	};

	class RecipientRowStream : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_cVersion;
		blockT<DWORD> m_cRowCount;
		std::vector<ADRENTRYStruct> m_lpAdrEntry;
	};
}