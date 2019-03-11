#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXOMSG].pdf
	struct ResponseLevel
	{
		blockT<bool> DeltaCode;
		blockT<DWORD> TimeDelta;
		blockT<BYTE> Random;
		blockT<BYTE> Level;

		ResponseLevel(std::shared_ptr<binaryParser> parser);
	};

	class ConversationIndex : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<BYTE> m_UnnamedByte;
		blockT<FILETIME> m_ftCurrent;
		blockT<GUID> m_guid;
		ULONG m_ulResponseLevels = 0;
		std::vector<std::shared_ptr<ResponseLevel>> m_lpResponseLevels;
	};
} // namespace smartview