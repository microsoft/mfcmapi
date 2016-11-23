#include "stdafx.h"
#include "ConversationIndex.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"

ConversationIndex::ConversationIndex()
{
	m_UnnamedByte = 0;
	m_ftCurrent = { 0 };
	m_guid = { 0 };
	m_ulResponseLevels = 0;
}

void ConversationIndex::Parse()
{
	m_UnnamedByte = m_Parser.Get<BYTE>();
	auto b1 = m_Parser.Get<BYTE>();
	auto b2 = m_Parser.Get<BYTE>();
	auto b3 = m_Parser.Get<BYTE>();
	// Encoding of the file time drops the high byte, which is always 1
	// So we add it back to get a time which makes more sense
	m_ftCurrent.dwHighDateTime = 1 << 24 | b1 << 16 | b2 << 8 | b3;

	b1 = m_Parser.Get<BYTE>();
	b2 = m_Parser.Get<BYTE>();
	m_ftCurrent.dwLowDateTime = b1 << 24 | b2 << 16;

	b1 = m_Parser.Get<BYTE>();
	b2 = m_Parser.Get<BYTE>();
	b3 = m_Parser.Get<BYTE>();
	auto b4 = m_Parser.Get<BYTE>();
	m_guid.Data1 = b1 << 24 | b2 << 16 | b3 << 8 | b4;

	b1 = m_Parser.Get<BYTE>();
	b2 = m_Parser.Get<BYTE>();
	m_guid.Data2 = static_cast<unsigned short>(b1 << 8 | b2);

	b1 = m_Parser.Get<BYTE>();
	b2 = m_Parser.Get<BYTE>();
	m_guid.Data3 = static_cast<unsigned short>(b1 << 8 | b2);

	for (auto i = 0 ; i < sizeof m_guid.Data4;i++)
	{
		m_guid.Data4[i] = m_Parser.Get<BYTE>();
	}

	if (m_Parser.RemainingBytes() > 0)
	{
		m_ulResponseLevels = static_cast<ULONG>(m_Parser.RemainingBytes()) / 5; // Response levels consume 5 bytes each
	}

	if (m_ulResponseLevels && m_ulResponseLevels < _MaxEntriesSmall)
	{
		for (ULONG i = 0; i < m_ulResponseLevels; i++)
		{
			ResponseLevel responseLevel;
			b1 = m_Parser.Get<BYTE>();
			b2 = m_Parser.Get<BYTE>();
			b3 = m_Parser.Get<BYTE>();
			b4 = m_Parser.Get<BYTE>();
			responseLevel.TimeDelta = b1 << 24 | b2 << 16 | b3 << 8 | b4;
			responseLevel.DeltaCode = false;
			if (responseLevel.TimeDelta & 0x80000000)
			{
				responseLevel.TimeDelta = responseLevel.TimeDelta & ~0x80000000;
				responseLevel.DeltaCode = true;
			}

			b1 = m_Parser.Get<BYTE>();
			responseLevel.Random = static_cast<BYTE>(b1 >> 4);
			responseLevel.Level = static_cast<BYTE>(b1 & 0xf);

			m_lpResponseLevels.push_back(responseLevel);
		}
	}
}

_Check_return_ wstring ConversationIndex::ToStringInternal()
{
	wstring szConversationIndex;

	wstring PropString;
	wstring AltPropString;
	FileTimeToString(m_ftCurrent, PropString, AltPropString);
	auto szGUID = GUIDToString(&m_guid);
	szConversationIndex = formatmessage(IDS_CONVERSATIONINDEXHEADER,
		m_UnnamedByte,
		m_ftCurrent.dwLowDateTime,
		m_ftCurrent.dwHighDateTime,
		PropString.c_str(),
		szGUID.c_str());

	if (m_lpResponseLevels.size())
	{
		for (ULONG i = 0; i < m_lpResponseLevels.size(); i++)
		{
			szConversationIndex += formatmessage(IDS_CONVERSATIONINDEXRESPONSELEVEL,
				i, m_lpResponseLevels[i].DeltaCode,
				m_lpResponseLevels[i].TimeDelta,
				m_lpResponseLevels[i].Random,
				m_lpResponseLevels[i].Level);
		}
	}

	return szConversationIndex;
}