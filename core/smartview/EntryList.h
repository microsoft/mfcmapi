#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/EntryIdStruct.h>

namespace smartview
{
	struct EntryListEntryStruct
	{
		std::shared_ptr<blockT<DWORD>> EntryLength = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> EntryLengthPad = emptyT<DWORD>();
		EntryIdStruct EntryId;

		EntryListEntryStruct(std::shared_ptr<binaryParser> parser);
	};

	class EntryList : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::shared_ptr<blockT<DWORD>> m_EntryCount = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_Pad = emptyT<DWORD>();

		std::vector<std::shared_ptr<EntryListEntryStruct>> m_Entry;
	};
} // namespace smartview