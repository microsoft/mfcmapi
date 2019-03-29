#include <core/stdafx.h>
#include <core/smartview/ConversationIndex.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>

namespace smartview
{
	ResponseLevel::ResponseLevel(const std::shared_ptr<binaryParser>& parser)
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

	void ConversationIndex::parse()
	{
		m_UnnamedByte = blockT<BYTE>::parse(m_Parser);
		const auto h1 = blockT<BYTE>::parse(m_Parser);
		const auto h2 = blockT<BYTE>::parse(m_Parser);
		const auto h3 = blockT<BYTE>::parse(m_Parser);

		// Encoding of the file time drops the high byte, which is always 1
		// So we add it back to get a time which makes more sense
		auto ft = FILETIME{};
		ft.dwHighDateTime = 1 << 24 | *h1 << 16 | *h2 << 8 | *h3;

		const auto l1 = blockT<BYTE>::parse(m_Parser);
		const auto l2 = blockT<BYTE>::parse(m_Parser);
		ft.dwLowDateTime = *l1 << 24 | *l2 << 16;

		m_ftCurrent = blockT<FILETIME>::create(
			ft, h1->getSize() + h2->getSize() + h3->getSize() + l1->getSize() + l2->getSize(), h1->getOffset());

		auto guid = GUID{};
		const auto g1 = blockT<BYTE>::parse(m_Parser);
		const auto g2 = blockT<BYTE>::parse(m_Parser);
		const auto g3 = blockT<BYTE>::parse(m_Parser);
		const auto g4 = blockT<BYTE>::parse(m_Parser);
		guid.Data1 = *g1 << 24 | *g2 << 16 | *g3 << 8 | *g4;

		const auto g5 = blockT<BYTE>::parse(m_Parser);
		const auto g6 = blockT<BYTE>::parse(m_Parser);
		guid.Data2 = static_cast<unsigned short>(*g5 << 8 | *g6);

		const auto g7 = blockT<BYTE>::parse(m_Parser);
		const auto g8 = blockT<BYTE>::parse(m_Parser);
		guid.Data3 = static_cast<unsigned short>(*g7 << 8 | *g8);

		auto size = g1->getSize() + g2->getSize() + g3->getSize() + g4->getSize() + g5->getSize() + g6->getSize() +
					g7->getSize() + g8->getSize();
		for (auto& i : guid.Data4)
		{
			const auto d = blockT<BYTE>::parse(m_Parser);
			i = *d;
			size += d->getSize();
		}

		m_guid = blockT<GUID>::create(guid, size, g1->getOffset());
		auto ulResponseLevels = ULONG{};
		if (m_Parser->getSize() > 0)
		{
			ulResponseLevels = static_cast<ULONG>(m_Parser->getSize()) / 5; // Response levels consume 5 bytes each
		}

		if (ulResponseLevels && ulResponseLevels < _MaxEntriesSmall)
		{
			m_lpResponseLevels.reserve(ulResponseLevels);
			for (ULONG i = 0; i < ulResponseLevels; i++)
			{
				m_lpResponseLevels.emplace_back(std::make_shared<ResponseLevel>(m_Parser));
			}
		}
	}

	void ConversationIndex::parseBlocks()
	{
		setRoot(L"Conversation Index: \r\n");

		std::wstring PropString;
		std::wstring AltPropString;
		strings::FileTimeToString(*m_ftCurrent, PropString, AltPropString);
		addChild(m_UnnamedByte, L"Unnamed byte = 0x%1!02X! = %1!d!\r\n", m_UnnamedByte->getData());
		addChild(
			m_ftCurrent,
			L"Current FILETIME: (Low = 0x%1!08X!, High = 0x%2!08X!) = %3!ws!\r\n",
			m_ftCurrent->getData().dwLowDateTime,
			m_ftCurrent->getData().dwHighDateTime,
			PropString.c_str());
		addChild(m_guid, L"GUID = %1!ws!", guid::GUIDToString(*m_guid).c_str());

		if (!m_lpResponseLevels.empty())
		{
			auto i = 0;
			for (const auto& responseLevel : m_lpResponseLevels)
			{
				terminateBlock();
				addChild(
					responseLevel->DeltaCode,
					L"ResponseLevel[%1!d!].DeltaCode = %2!d!\r\n",
					i,
					responseLevel->DeltaCode->getData());
				addChild(
					responseLevel->TimeDelta,
					L"ResponseLevel[%1!d!].TimeDelta = 0x%2!08X! = %2!d!\r\n",
					i,
					responseLevel->TimeDelta->getData());
				addChild(
					responseLevel->Random,
					L"ResponseLevel[%1!d!].Random = 0x%2!02X! = %2!d!\r\n",
					i,
					responseLevel->Random->getData());
				addChild(
					responseLevel->Level,
					L"ResponseLevel[%1!d!].ResponseLevel = 0x%2!02X! = %2!d!",
					i,
					responseLevel->Level->getData());

				i++;
			}
		}
	}
} // namespace smartview