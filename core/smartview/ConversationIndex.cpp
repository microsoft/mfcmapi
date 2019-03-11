#include <core/stdafx.h>
#include <core/smartview/ConversationIndex.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>

namespace smartview
{
	ResponseLevel::ResponseLevel(std::shared_ptr<binaryParser> parser)
	{
		auto& r1 = blockT<BYTE>(parser);
		auto& r2 = blockT<BYTE>(parser);
		auto& r3 = blockT<BYTE>(parser);
		auto& r4 = blockT<BYTE>(parser);
		TimeDelta.setOffset(r1.getOffset());
		TimeDelta.setSize(r1.getSize() + r2.getSize() + r3.getSize() + r4.getSize());
		TimeDelta.setData(r1 << 24 | r2 << 16 | r3 << 8 | r4);

		DeltaCode.setOffset(TimeDelta.getOffset());
		DeltaCode.setSize(TimeDelta.getSize());
		DeltaCode.setData(false);
		if (TimeDelta & 0x80000000)
		{
			TimeDelta.setData(TimeDelta & ~0x80000000);
			DeltaCode.setData(true);
		}

		auto& r5 = blockT<BYTE>(parser);
		Random.setOffset(r5.getOffset());
		Random.setSize(r5.getSize());
		Random.setData(static_cast<BYTE>(r5 >> 4));

		Level.setOffset(r5.getOffset());
		Level.setSize(r5.getSize());
		Level.setData(static_cast<BYTE>(r5 & 0xf));
	}

	void ConversationIndex::Parse()
	{
		m_UnnamedByte.parse<BYTE>(m_Parser);
		auto& h1 = blockT<BYTE>(m_Parser);
		auto& h2 = blockT<BYTE>(m_Parser);
		auto& h3 = blockT<BYTE>(m_Parser);

		// Encoding of the file time drops the high byte, which is always 1
		// So we add it back to get a time which makes more sense
		auto ft = FILETIME{};
		ft.dwHighDateTime = 1 << 24 | h1 << 16 | h2 << 8 | h3;

		auto& l1 = blockT<BYTE>(m_Parser);
		auto& l2 = blockT<BYTE>(m_Parser);
		ft.dwLowDateTime = l1 << 24 | l2 << 16;

		m_ftCurrent.setOffset(h1.getOffset());
		m_ftCurrent.setSize(h1.getSize() + h2.getSize() + h3.getSize() + l1.getSize() + l2.getSize());
		m_ftCurrent.setData(ft);

		auto guid = GUID{};
		auto& g1 = blockT<BYTE>(m_Parser);
		auto& g2 = blockT<BYTE>(m_Parser);
		auto& g3 = blockT<BYTE>(m_Parser);
		auto& g4 = blockT<BYTE>(m_Parser);
		guid.Data1 = g1 << 24 | g2 << 16 | g3 << 8 | g4;

		auto& g5 = blockT<BYTE>(m_Parser);
		auto& g6 = blockT<BYTE>(m_Parser);
		guid.Data2 = static_cast<unsigned short>(g5 << 8 | g6);

		auto& g7 = blockT<BYTE>(m_Parser);
		auto& g8 = blockT<BYTE>(m_Parser);
		guid.Data3 = static_cast<unsigned short>(g7 << 8 | g8);

		auto size = g1.getSize() + g2.getSize() + g3.getSize() + g4.getSize() + g5.getSize() + g6.getSize() +
					g7.getSize() + g8.getSize();
		for (auto& i : guid.Data4)
		{
			auto& d = blockT<BYTE>(m_Parser);
			i = d;
			size += d.getSize();
		}

		m_guid.setOffset(g1.getOffset());
		m_guid.setSize(size);
		m_guid.setData(guid);

		if (m_Parser->RemainingBytes() > 0)
		{
			m_ulResponseLevels =
				static_cast<ULONG>(m_Parser->RemainingBytes()) / 5; // Response levels consume 5 bytes each
		}

		if (m_ulResponseLevels && m_ulResponseLevels < _MaxEntriesSmall)
		{
			m_lpResponseLevels.reserve(m_ulResponseLevels);
			for (ULONG i = 0; i < m_ulResponseLevels; i++)
			{
				m_lpResponseLevels.emplace_back(std::make_shared<ResponseLevel>(m_Parser));
			}
		}
	}

	void ConversationIndex::ParseBlocks()
	{
		setRoot(L"Conversation Index: \r\n");

		std::wstring PropString;
		std::wstring AltPropString;
		strings::FileTimeToString(m_ftCurrent, PropString, AltPropString);
		addBlock(m_UnnamedByte, L"Unnamed byte = 0x%1!02X! = %1!d!\r\n", m_UnnamedByte.getData());
		addBlock(
			m_ftCurrent,
			L"Current FILETIME: (Low = 0x%1!08X!, High = 0x%2!08X!) = %3!ws!\r\n",
			m_ftCurrent.getData().dwLowDateTime,
			m_ftCurrent.getData().dwHighDateTime,
			PropString.c_str());
		addBlock(m_guid, L"GUID = %1!ws!", guid::GUIDToString(m_guid).c_str());

		if (!m_lpResponseLevels.empty())
		{
			auto i = 0;
			for (const auto& responseLevel : m_lpResponseLevels)
			{
				terminateBlock();
				addBlock(
					responseLevel->DeltaCode,
					L"ResponseLevel[%1!d!].DeltaCode = %2!d!\r\n",
					i,
					responseLevel->DeltaCode.getData());
				addBlock(
					responseLevel->TimeDelta,
					L"ResponseLevel[%1!d!].TimeDelta = 0x%2!08X! = %2!d!\r\n",
					i,
					responseLevel->TimeDelta.getData());
				addBlock(
					responseLevel->Random,
					L"ResponseLevel[%1!d!].Random = 0x%2!02X! = %2!d!\r\n",
					i,
					responseLevel->Random.getData());
				addBlock(
					responseLevel->Level,
					L"ResponseLevel[%1!d!].ResponseLevel = 0x%2!02X! = %2!d!",
					i,
					responseLevel->Level.getData());
				i++;
			}
		}
	}
} // namespace smartview