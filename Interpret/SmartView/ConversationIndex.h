#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	// [MS-OXOMSG].pdf
	struct ResponseLevel
	{
		blockT<bool> DeltaCode;
		blockT<DWORD> TimeDelta;
		blockT<BYTE> Random;
		blockT<BYTE> Level;
	};

	class ConversationIndex : public SmartViewParser
	{
	public:
		ConversationIndex();

	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<BYTE> m_UnnamedByte;
		blockT<FILETIME> m_ftCurrent;
		blockT<GUID> m_guid;
		ULONG m_ulResponseLevels;
		std::vector<ResponseLevel> m_lpResponseLevels;
	};
}