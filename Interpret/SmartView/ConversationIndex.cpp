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
		ft.dwHighDateTime = 1 << 24 | h1.getData() << 16 | h2.getData() << 8 | h3.getData();

		auto l1 = m_Parser.GetBlock<BYTE>();
		auto l2 = m_Parser.GetBlock<BYTE>();
		ft.dwLowDateTime = l1.getData() << 24 | l2.getData() << 16;

		m_ftCurrent.setOffset(h1.getOffset());
		m_ftCurrent.setSize(h1.getSize() + h2.getSize() + h3.getSize() + l1.getSize() + l2.getSize());
		m_ftCurrent.setData(ft);

		auto guid = GUID{};
		auto g1 = m_Parser.GetBlock<BYTE>();
		auto g2 = m_Parser.GetBlock<BYTE>();
		auto g3 = m_Parser.GetBlock<BYTE>();
		auto g4 = m_Parser.GetBlock<BYTE>();
		guid.Data1 = g1.getData() << 24 | g2.getData() << 16 | g3.getData() << 8 | g4.getData();

		auto g5 = m_Parser.GetBlock<BYTE>();
		auto g6 = m_Parser.GetBlock<BYTE>();
		guid.Data2 = static_cast<unsigned short>(g5.getData() << 8 | g6.getData());

		auto g7 = m_Parser.GetBlock<BYTE>();
		auto g8 = m_Parser.GetBlock<BYTE>();
		guid.Data3 = static_cast<unsigned short>(g7.getData() << 8 | g8.getData());

		auto size = g1.getSize() + g2.getSize() + g3.getSize() + g4.getSize() + g5.getSize() + g6.getSize() +
					g7.getSize() + g8.getSize();
		for (auto& i : guid.Data4)
		{
			auto d = m_Parser.GetBlock<BYTE>();
			i = d.getData();
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
					r1.getData() << 24 | r2.getData() << 16 | r3.getData() << 8 | r4.getData());

				responseLevel.DeltaCode.setOffset(responseLevel.TimeDelta.getOffset());
				responseLevel.DeltaCode.setSize(responseLevel.TimeDelta.getSize());
				responseLevel.DeltaCode.setData(false);
				if (responseLevel.TimeDelta.getData() & 0x80000000)
				{
					responseLevel.TimeDelta.setData(responseLevel.TimeDelta.getData() & ~0x80000000);
					responseLevel.DeltaCode.setData(true);
				}

				auto r5 = m_Parser.GetBlock<BYTE>();
				responseLevel.Random.setOffset(r5.getOffset());
				responseLevel.Random.setSize(r5.getSize());
				responseLevel.Random.setData(static_cast<BYTE>(r5.getData() >> 4));

				responseLevel.Level.setOffset(r5.getOffset());
				responseLevel.Level.setSize(r5.getSize());
				responseLevel.Level.setData(static_cast<BYTE>(r5.getData() & 0xf));

				m_lpResponseLevels.push_back(responseLevel);
			}
		}
	}

	_Check_return_ std::wstring ConversationIndex::ToStringInternal()
	{
		std::wstring PropString;
		std::wstring AltPropString;
		strings::FileTimeToString(m_ftCurrent.getData(), PropString, AltPropString);
		auto szGUID = guid::GUIDToString(m_guid.getData());
		auto szConversationIndex = strings::formatmessage(
			IDS_CONVERSATIONINDEXHEADER,
			m_UnnamedByte.getData(),
			m_ftCurrent.getData().dwLowDateTime,
			m_ftCurrent.getData().dwHighDateTime,
			PropString.c_str(),
			szGUID.c_str());

		if (m_lpResponseLevels.size())
		{
			for (ULONG i = 0; i < m_lpResponseLevels.size(); i++)
			{
				szConversationIndex += strings::formatmessage(
					IDS_CONVERSATIONINDEXRESPONSELEVEL,
					i,
					m_lpResponseLevels[i].DeltaCode.getData(),
					m_lpResponseLevels[i].TimeDelta.getData(),
					m_lpResponseLevels[i].Random.getData(),
					m_lpResponseLevels[i].Level.getData());
			}
		}

		return szConversationIndex;
	}
}