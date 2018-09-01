#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/PropertyStruct.h>

namespace smartview
{
	struct ADRENTRYStruct
	{
		blockT<DWORD> ulReserved1;
		blockT<DWORD> cValues;
		PropertyStruct rgPropVals;
	} ;

	class RecipientRowStream : public SmartViewParser
	{
	public:
		RecipientRowStream();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		blockT<DWORD> m_cVersion;
		blockT<DWORD> m_cRowCount;
		std::vector<ADRENTRYStruct> m_lpAdrEntry;
	};
}