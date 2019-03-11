#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct SRowStruct
	{
		blockT<DWORD> cValues;
		PropertiesStruct lpProps;

		SRowStruct(std::shared_ptr<binaryParser> parser);
	};

	class NickNameCache : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockBytes m_Metadata1; // 4 bytes
		blockT<DWORD> m_ulMajorVersion;
		blockT<DWORD> m_ulMinorVersion;
		blockT<DWORD> m_cRowCount;
		std::vector<std::shared_ptr<SRowStruct>> m_lpRows;
		blockT<DWORD> m_cbEI;
		blockBytes m_lpbEI;
		blockBytes m_Metadata2; // 8 bytes
	};
} // namespace smartview