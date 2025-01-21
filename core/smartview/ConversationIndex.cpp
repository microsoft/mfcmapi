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
		unsigned int td = *TimeDelta;
		if (*DeltaCode)
		{
			ft.dwHighDateTime |= ((td >> 24) & 0x7F) << 15 | ((td >> 16) & 0xFF) << 7 | ((td >> 8) & 0xFF) >> 1;
			ft.dwLowDateTime = (static_cast<DWORD>((td >> 8) & 0xFF) << 31) | (static_cast<DWORD>(td & 0xFF) << 23);
		}
		else
		{
			ft.dwHighDateTime |= ((td >> 24) & 0x7F) << 10 | ((td >> 16) & 0xFF) << 2 | ((td >> 8) & 0xFF) >> 6;
			ft.dwLowDateTime = (static_cast<DWORD>((td >> 8) & 0xFF) << 26) | (static_cast<DWORD>(td & 0xFF) << 18);
		}

		ft = AddFileTime(currentFileTime, ft);

		std::wstring PropString;
		std::wstring AltPropString;
		strings::FileTimeToString(ft, PropString, AltPropString);
		const auto ftBlock = blockT<FILETIME>::create(ft, TimeDelta->getSize(), TimeDelta->getOffset());
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

		std::wstring propString;
		std::wstring altPropString;
		strings::FileTimeToString(*currentFileTime, propString, altPropString);

		addChild(currentFileTime, L"Current FILETIME: %1!ws!", propString.c_str());
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