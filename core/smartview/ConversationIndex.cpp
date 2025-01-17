#include <core/stdafx.h>
#include <core/smartview/ConversationIndex.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>

namespace smartview
{
	void ResponseLevel::parse()
	{
		TimeDelta = blockT<DWORD>::parse(parser);
		TimeDelta->setData(std::byteswap(TimeDelta->getData()));

		DeltaCode = blockT<bool>::create(false, 1, TimeDelta->getOffset());
		if (*TimeDelta & 0x80000000)
		{
			TimeDelta->setData(*TimeDelta & ~0x80000000);
			DeltaCode->setData(true);
		}

		const auto r5 = blockT<BYTE>::parse(parser);
		Random = blockT<BYTE>::create(static_cast<BYTE>(*r5 >> 4), r5->getSize(), r5->getOffset());
		Level = blockT<BYTE>::create(static_cast<BYTE>(*r5 & 0xf), r5->getSize(), r5->getOffset());
	}

	_Check_return_ FILETIME AddFileTime(const FILETIME& ft, const FILETIME& dwDelta)
	{
		auto ftResult = FILETIME{};
		auto li = LARGE_INTEGER{};
		li.LowPart = ft.dwLowDateTime;
		li.HighPart = ft.dwHighDateTime;
		auto liDelta = LARGE_INTEGER{};
		liDelta.LowPart = dwDelta.dwLowDateTime;
		liDelta.HighPart = dwDelta.dwHighDateTime;
		li.QuadPart += liDelta.QuadPart;
		ftResult.dwLowDateTime = li.LowPart;
		ftResult.dwHighDateTime = li.HighPart;
		return ftResult;
	}

	void ResponseLevel::parseBlocks()
	{
		addChild(DeltaCode, L"DeltaCode = %1!d!", DeltaCode->getData());
		auto ft = FILETIME{};
		unsigned char b1 = (TimeDelta->getData() >> 24) & 0x7F;
		unsigned char b2 = (TimeDelta->getData() >> 16) & 0xFF;
		unsigned char b3 = (TimeDelta->getData() >> 8) & 0xFF;
		unsigned char b4 = TimeDelta->getData() & 0xFF;
		if (*DeltaCode)
		{
			ft.dwHighDateTime |= b1 << 15;
			ft.dwHighDateTime |= b2 << 7;
			ft.dwHighDateTime |= b3 >> 1;
			ft.dwLowDateTime = (((DWORD) (b3)) << 31) | (((DWORD) (b4)) << 23);
		}
		else
		{
			ft.dwHighDateTime |= b1 << 10;
			ft.dwHighDateTime |= b2 << 2;
			ft.dwHighDateTime |= b3 >> 6;
			ft.dwLowDateTime = (((DWORD) (b3)) << 26) | (((DWORD) (b4)) << 18);
		}

		ft = AddFileTime(currentFileTime, ft);

		std::wstring PropString;
		std::wstring AltPropString;
		strings::FileTimeToString(ft, PropString, AltPropString);
		const auto ftBlock = blockT<FILETIME>::create(ft, TimeDelta->getSize(), 
			TimeDelta->getOffset());
		addChild(ftBlock, L"Time = %1!ws!", PropString.c_str());

		ftBlock->addChild(TimeDelta, L"TimeDelta = 0x%1!08X! = %1!d!", TimeDelta->getData());
		addChild(Random, L"Random = 0x%1!02X! = %1!d!", Random->getData());
		addChild(Level, L"ResponseLevel = 0x%1!02X! = %1!d!", Level->getData());
	}

	void ConversationIndex::parse()
	{
		dwHighDateTime = blockT<DWORD>::parse(parser);
		dwHighDateTime->setData(std::byteswap(dwHighDateTime->getData()));
		dwLowDateTime = blockT<WORD>::parse(parser);
		dwLowDateTime->setData(std::byteswap(dwLowDateTime->getData()));
		currentFileTime = blockT<FILETIME>::create(
			FILETIME{*dwLowDateTime, *dwHighDateTime},
			dwHighDateTime->getSize() + dwLowDateTime->getSize(),
			dwHighDateTime->getOffset());

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
				auto responseLevel = std::make_shared<smartview::ResponseLevel>(currentFileTime->getData());
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

		addChild(currentFileTime, L"Current FILETIME: %1!ws!", PropString.c_str());
		currentFileTime->addChild(dwLowDateTime, L"LowDateTime = 0x%1!04X!", dwLowDateTime->getData());
		currentFileTime->addChild(dwHighDateTime, L"HighDateTime = 0x%1!08X!", dwHighDateTime->getData());
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