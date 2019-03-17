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

		std::shared_ptr<blockBytes> m_Metadata1 = emptyBB(); // 4 bytes
		blockT<DWORD> m_ulMajorVersion;
		blockT<DWORD> m_ulMinorVersion;
		blockT<DWORD> m_cRowCount;
		std::vector<std::shared_ptr<SRowStruct>> m_lpRows;
		blockT<DWORD> m_cbEI;
		std::shared_ptr<blockBytes> m_lpbEI = emptyBB();
		std::shared_ptr<blockBytes> m_Metadata2 = emptyBB(); // 8 bytes
	};
} // namespace smartview