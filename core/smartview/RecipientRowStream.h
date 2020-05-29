#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct ADRENTRYStruct
	{
		std::shared_ptr<blockT<DWORD>> ulReserved1 = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> cValues = emptyT<DWORD>();
		PropertiesStruct rgPropVals;

		ADRENTRYStruct(const std::shared_ptr<binaryParser>& parser);
	};

	class RecipientRowStream : public smartViewParser
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> m_cVersion = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_cRowCount = emptyT<DWORD>();
		std::vector<std::shared_ptr<ADRENTRYStruct>> m_lpAdrEntry;
	};
} // namespace smartview