#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXOMSG].pdf
	struct ResponseLevel
	{
		std::shared_ptr<blockT<bool>> DeltaCode = emptyT<bool>();
		std::shared_ptr<blockT<DWORD>> TimeDelta = emptyT<DWORD>();
		std::shared_ptr<blockT<BYTE>> Random = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> Level = emptyT<BYTE>();

		ResponseLevel(const std::shared_ptr<binaryParser>& parser);
	};

	class ConversationIndex : public smartViewParser
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<BYTE>> m_UnnamedByte = emptyT<BYTE>();
		std::shared_ptr<blockT<FILETIME>> m_ftCurrent = emptyT<FILETIME>();
		std::shared_ptr<blockT<GUID>> m_guid = emptyT<GUID>();
		std::vector<std::shared_ptr<ResponseLevel>> m_lpResponseLevels;
	};
} // namespace smartview