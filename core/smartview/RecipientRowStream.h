#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/PropertiesStruct.h>

namespace smartview
{
	struct ADRENTRYStruct
	{
		blockT<DWORD> ulReserved1;
		blockT<DWORD> cValues;
		PropertiesStruct rgPropVals;

		ADRENTRYStruct(std::shared_ptr<binaryParser> parser);
	};

	class RecipientRowStream : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_cVersion;
		blockT<DWORD> m_cRowCount;
		std::vector<std::shared_ptr<ADRENTRYStruct>> m_lpAdrEntry;
	};
} // namespace smartview