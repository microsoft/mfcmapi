#pragma once
#include "SmartViewParser.h"

namespace smartview
{
	class NickNameCache : public SmartViewParser
	{
	public:
		NickNameCache();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		_Check_return_ LPSPropValue NickNameBinToSPropValue(DWORD dwPropCount);

		std::vector<BYTE> m_Metadata1; // 4 bytes
		ULONG m_ulMajorVersion;
		ULONG m_ulMinorVersion;
		DWORD m_cRowCount;
		LPSRow m_lpRows;
		ULONG m_cbEI;
		std::vector<BYTE> m_lpbEI;
		std::vector<BYTE> m_Metadata2; // 8 bytes
	};
}