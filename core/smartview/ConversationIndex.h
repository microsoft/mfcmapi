#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXOMSG] 2.2.1.3 PidTagConversationIndex Property
	// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxomsg/9e994fbb-b839-495f-84e3-2c8c02c7dd9b

	class ResponseLevel : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<bool>> DeltaCode = emptyT<bool>();
		std::shared_ptr<blockT<DWORD>> TimeDelta = emptyT<DWORD>();
		std::shared_ptr<blockT<BYTE>> Random = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> Level = emptyT<BYTE>();
	};

	class ConversationIndex : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<BYTE>> reserved = emptyT<BYTE>();
		std::shared_ptr<blockT<FILETIME>> currentFileTime = emptyT<FILETIME>();
		std::shared_ptr<blockT<GUID>> threadGuid = emptyT<GUID>();
		std::vector<std::shared_ptr<ResponseLevel>> responseLevels;
	};
} // namespace smartview