#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/PropertiesStruct.h>

namespace smartview
{
	struct SRowStruct
	{
		blockT<DWORD> cValues;
		PropertiesStruct lpProps;
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
		std::vector<SRowStruct> m_lpRows;
		blockT<DWORD> m_cbEI;
		blockBytes m_lpbEI;
		blockBytes m_Metadata2; // 8 bytes
	};
}