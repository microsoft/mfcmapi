#include <core/stdafx.h>
#include <core/smartview/ConversationIndex.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>

namespace smartview
{
	void ResponseLevel::parse()
	{
		const auto r1 = blockT<BYTE>::parse(parser);
		const auto r2 = blockT<BYTE>::parse(parser);
		const auto r3 = blockT<BYTE>::parse(parser);
		const auto r4 = blockT<BYTE>::parse(parser);

		TimeDelta = blockT<DWORD>::create(
			static_cast<DWORD>(*r1 << 24 | *r2 << 16 | *r3 << 8 | *r4),
			r1->getSize() + r2->getSize() + r3->getSize() + r4->getSize(),
			r1->getOffset());

		DeltaCode = blockT<bool>::create(false, TimeDelta->getSize(), TimeDelta->getOffset());
		if (*TimeDelta & 0x80000000)
		{
			TimeDelta->setData(*TimeDelta & ~0x80000000);
			DeltaCode->setData(true);
		}

		const auto r5 = blockT<BYTE>::parse(parser);
		Random = blockT<BYTE>::create(static_cast<BYTE>(*r5 >> 4), r5->getSize(), r5->getOffset());
		Level = blockT<BYTE>::create(static_cast<BYTE>(*r5 & 0xf), r5->getSize(), r5->getOffset());
	}

	void ResponseLevel::parseBlocks()
	{
		addChild(DeltaCode, L"DeltaCode = %1!d!", DeltaCode->getData());
		addChild(TimeDelta, L"TimeDelta = 0x%1!08X! = %1!d!", TimeDelta->getData());
		addChild(Random, L"Random = 0x%1!02X! = %1!d!", Random->getData());
		addChild(Level, L"ResponseLevel = 0x%1!02X! = %1!d!", Level->getData());
	}

	void ConversationIndex::parse()
	{
		reserved = blockT<BYTE>::parse(parser);
		const auto h1 = blockT<BYTE>::parse(parser);
		const auto h2 = blockT<BYTE>::parse(parser);
		const auto h3 = blockT<BYTE>::parse(parser);

		// Encoding of the file time drops the high byte, which is always 1
		// So we add it back to get a time which makes more sense
		auto ft = FILETIME{};
		ft.dwHighDateTime = 1 << 24 | *h1 << 16 | *h2 << 8 | *h3;

		const auto l1 = blockT<BYTE>::parse(parser);
		const auto l2 = blockT<BYTE>::parse(parser);
		ft.dwLowDateTime = *l1 << 24 | *l2 << 16;

		currentFileTime = blockT<FILETIME>::create(
			ft, h1->getSize() + h2->getSize() + h3->getSize() + l1->getSize() + l2->getSize(), h1->getOffset());

		threadGuid = blockT<GUID>::parse(parser);
		auto ulResponseLevels = ULONG{};
		if (parser->getSize() > 0)
		{
			ulResponseLevels = static_cast<ULONG>(parser->getSize()) / 5; // Response levels consume 5 bytes each
		}

		if (ulResponseLevels && ulResponseLevels < _MaxEntriesSmall)
		{
			responseLevels.reserve(ulResponseLevels);
			for (ULONG i = 0; i < ulResponseLevels; i++)
			{
				auto responseLevel = std::make_shared<smartview::ResponseLevel>();
				responseLevel->block::parse(parser, false);
				if (!responseLevel->isSet()) break;
				responseLevels.emplace_back(responseLevel);
			}
		}
	}

	void ConversationIndex::parseBlocks()
	{
		setText(L"Conversation Index");

		std::wstring PropString;
		std::wstring AltPropString;
		strings::FileTimeToString(*currentFileTime, PropString, AltPropString);
		addChild(reserved, L"Unnamed byte = 0x%1!02X! = %1!d!", reserved->getData());
		addChild(
			currentFileTime,
			L"Current FILETIME: (Low = 0x%1!08X!, High = 0x%2!08X!) = %3!ws!",
			currentFileTime->getData().dwLowDateTime,
			currentFileTime->getData().dwHighDateTime,
			PropString.c_str());
		addChild(threadGuid, L"GUID = %1!ws!", guid::GUIDToString(*threadGuid).c_str());

		if (!responseLevels.empty())
		{
			auto i = 0;
			for (const auto& responseLevel : responseLevels)
			{
				addChild(responseLevel, L"ResponseLevel[%1!d!]", i);
				i++;
			}
		}
	}
} // namespace smartview