#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct SRowStruct
	{
		std::shared_ptr<blockT<DWORD>> cValues = emptyT<DWORD>();
		std::shared_ptr<PropertiesStruct> lpProps;

		SRowStruct(const std::shared_ptr<binaryParser>& parser);
	};

	class NickNameCache : public smartViewParser
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockBytes> m_Metadata1 = emptyBB(); // 4 bytes
		std::shared_ptr<blockT<DWORD>> m_ulMajorVersion = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_ulMinorVersion = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_cRowCount = emptyT<DWORD>();
		std::vector<std::shared_ptr<SRowStruct>> m_lpRows;
		std::shared_ptr<blockT<DWORD>> m_cbEI = emptyT<DWORD>();
		std::shared_ptr<blockBytes> m_lpbEI = emptyBB();
		std::shared_ptr<blockBytes> m_Metadata2 = emptyBB(); // 8 bytes
	};
} // namespace smartview