#include <StdAfx.h>
#include <Interpret/SmartView/ConversationIndex.h>
#include <Interpret/Guids.h>

namespace smartview
{
	ConversationIndex::ConversationIndex() {}

	void ConversationIndex::Parse()
	{
		m_UnnamedByte = m_Parser.GetBlock<BYTE>();
		auto h1 = m_Parser.GetBlock<BYTE>();
		auto h2 = m_Parser.GetBlock<BYTE>();
		auto h3 = m_Parser.GetBlock<BYTE>();

		// Encoding of the file time drops the high byte, which is always 1
		// So we add it back to get a time which makes more sense
		auto ft = FILETIME{};
		ft.dwHighDateTime = 1 << 24 | h1 << 16 | h2 << 8 | h3;

		auto l1 = m_Parser.GetBlock<BYTE>();
		auto l2 = m_Parser.GetBlock<BYTE>();
		ft.dwLowDateTime = l1 << 24 | l2 << 16;

		m_ftCurrent.setOffset(h1.getOffset());
		m_ftCurrent.setSize(h1.getSize() + h2.getSize() + h3.getSize() + l1.getSize() + l2.getSize());
		m_ftCurrent.setData(ft);

		auto guid = GUID{};
		auto g1 = m_Parser.GetBlock<BYTE>();
		auto g2 = m_Parser.GetBlock<BYTE>();
		auto g3 = m_Parser.GetBlock<BYTE>();
		auto g4 = m_Parser.GetBlock<BYTE>();
		guid.Data1 = g1 << 24 | g2 << 16 | g3 << 8 | g4;

		auto g5 = m_Parser.GetBlock<BYTE>();
		auto g6 = m_Parser.GetBlock<BYTE>();
		guid.Data2 = static_cast<unsigned short>(g5 << 8 | g6);

		auto g7 = m_Parser.GetBlock<BYTE>();
		auto g8 = m_Parser.GetBlock<BYTE>();
		guid.Data3 = static_cast<unsigned short>(g7 << 8 | g8);

		auto size = g1.getSize() + g2.getSize() + g3.getSize() + g4.getSize() + g5.getSize() + g6.getSize() +
					g7.getSize() + g8.getSize();
		for (auto& i : guid.Data4)
		{
			auto d = m_Parser.GetBlock<BYTE>();
			i = d;
			size += d.getSize();
		}

		m_guid.setOffset(g1.getOffset());
		m_guid.setSize(size);
		m_guid.setData(guid);

		if (m_Parser.RemainingBytes() > 0)
		{
			m_ulResponseLevels =
				static_cast<ULONG>(m_Parser.RemainingBytes()) / 5; // Response levels consume 5 bytes each
		}

		if (m_ulResponseLevels && m_ulResponseLevels < _MaxEntriesSmall)
		{
			for (ULONG i = 0; i < m_ulResponseLevels; i++)
			{
				ResponseLevel responseLevel{};
				auto r1 = m_Parser.GetBlock<BYTE>();
				auto r2 = m_Parser.GetBlock<BYTE>();
				auto r3 = m_Parser.GetBlock<BYTE>();
				auto r4 = m_Parser.GetBlock<BYTE>();
				responseLevel.TimeDelta.setOffset(r1.getOffset());
				responseLevel.TimeDelta.setSize(r1.getSize() + r2.getSize() + r3.getSize() + r4.getSize());
				responseLevel.TimeDelta.setData(
					r1 << 24 | r2 << 16 | r3 << 8 | r4);

				responseLevel.DeltaCode.setOffset(responseLevel.TimeDelta.getOffset());
				responseLevel.DeltaCode.setSize(responseLevel.TimeDelta.getSize());
				responseLevel.DeltaCode.setData(false);
				if (responseLevel.TimeDelta & 0x80000000)
				{
					responseLevel.TimeDelta.setData(responseLevel.TimeDelta & ~0x80000000);
					responseLevel.DeltaCode.setData(true);
				}

				auto r5 = m_Parser.GetBlock<BYTE>();
				responseLevel.Random.setOffset(r5.getOffset());
				responseLevel.Random.setSize(r5.getSize());
				responseLevel.Random.setData(static_cast<BYTE>(r5 >> 4));

				responseLevel.Level.setOffset(r5.getOffset());
				responseLevel.Level.setSize(r5.getSize());
				responseLevel.Level.setData(static_cast<BYTE>(r5 & 0xf));

				m_lpResponseLevels.push_back(responseLevel);
			}
		}
	}

	void ConversationIndex::ParseBlocks()
	{
		addHeader(L"Conversation Index: \r\n");

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

		if (m_lpResponseLevels.size())
		{
			for (ULONG i = 0; i < m_lpResponseLevels.size(); i++)
			{
				addBlock(
					m_lpResponseLevels[i].DeltaCode,
					L"\r\nResponseLevel[%1!d!].DeltaCode = %2!d!",
					i,
					m_lpResponseLevels[i].DeltaCode.getData());
				addBlock(
					m_lpResponseLevels[i].TimeDelta,
					L"\r\nResponseLevel[%1!d!].TimeDelta = 0x%2!08X! = %2!d!",
					i,
					m_lpResponseLevels[i].TimeDelta.getData());
				addBlock(
					m_lpResponseLevels[i].Random,
					L"\r\nResponseLevel[%1!d!].Random = 0x%2!02X! = %2!d!",
					i,
					m_lpResponseLevels[i].Random.getData());
				addBlock(
					m_lpResponseLevels[i].Level,
					L"\r\nResponseLevel[%1!d!].ResponseLevel = 0x%2!02X! = %2!d!",
					i,
					m_lpResponseLevels[i].Level.getData());
			}
		}
	}
}