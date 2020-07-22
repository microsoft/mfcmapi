#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXOMSG].pdf
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

		std::shared_ptr<blockT<BYTE>> m_UnnamedByte = emptyT<BYTE>();
		std::shared_ptr<blockT<FILETIME>> m_ftCurrent = emptyT<FILETIME>();
		std::shared_ptr<blockT<GUID>> m_guid = emptyT<GUID>();
		std::vector<std::shared_ptr<ResponseLevel>> m_lpResponseLevels;
	};
} // namespace smartview