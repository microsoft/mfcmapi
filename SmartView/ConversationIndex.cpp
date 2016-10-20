#include "stdafx.h"
#include "ConversationIndex.h"
#include "String.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"

ConversationIndex::ConversationIndex()
{
	m_UnnamedByte = 0;
	m_ftCurrent = { 0 };
	m_guid = { 0 };
	m_ulResponseLevels = 0;
	m_lpResponseLevels = nullptr;
}

ConversationIndex::~ConversationIndex()
{
	delete[] m_lpResponseLevels;
}

void ConversationIndex::Parse()
{
	m_Parser.GetBYTE(&m_UnnamedByte);
	BYTE b1 = NULL;
	BYTE b2 = NULL;
	BYTE b3 = NULL;
	BYTE b4 = NULL;
	m_Parser.GetBYTE(&b1);
	m_Parser.GetBYTE(&b2);
	m_Parser.GetBYTE(&b3);
	m_ftCurrent.dwHighDateTime = b1 << 16 | b2 << 8 | b3;
	m_Parser.GetBYTE(&b1);
	m_Parser.GetBYTE(&b2);
	m_ftCurrent.dwLowDateTime = b1 << 24 | b2 << 16;
	m_Parser.GetBYTE(&b1);
	m_Parser.GetBYTE(&b2);
	m_Parser.GetBYTE(&b3);
	m_Parser.GetBYTE(&b4);
	m_guid.Data1 = b1 << 24 | b2 << 16 | b3 << 8 | b4;
	m_Parser.GetBYTE(&b1);
	m_Parser.GetBYTE(&b2);
	m_guid.Data2 = static_cast<unsigned short>(b1 << 8 | b2);
	m_Parser.GetBYTE(&b1);
	m_Parser.GetBYTE(&b2);
	m_guid.Data3 = static_cast<unsigned short>(b1 << 8 | b2);
	m_Parser.GetBYTESNoAlloc(sizeof m_guid.Data4, sizeof m_guid.Data4, m_guid.Data4);

	if (m_Parser.RemainingBytes() > 0)
	{
		m_ulResponseLevels = static_cast<ULONG>(m_Parser.RemainingBytes()) / 5; // Response levels consume 5 bytes each
	}

	if (m_ulResponseLevels && m_ulResponseLevels < _MaxEntriesSmall)
		m_lpResponseLevels = new ResponseLevel[m_ulResponseLevels];

	if (m_lpResponseLevels)
	{
		memset(m_lpResponseLevels, 0, sizeof(ResponseLevel)* m_ulResponseLevels);
		for (ULONG i = 0; i < m_ulResponseLevels; i++)
		{
			m_Parser.GetBYTE(&b1);
			m_Parser.GetBYTE(&b2);
			m_Parser.GetBYTE(&b3);
			m_Parser.GetBYTE(&b4);
			m_lpResponseLevels[i].TimeDelta = b1 << 24 | b2 << 16 | b3 << 8 | b4;
			if (m_lpResponseLevels[i].TimeDelta & 0x80000000)
			{
				m_lpResponseLevels[i].TimeDelta = m_lpResponseLevels[i].TimeDelta & ~0x80000000;
				m_lpResponseLevels[i].DeltaCode = true;
			}

			m_Parser.GetBYTE(&b1);
			m_lpResponseLevels[i].Random = static_cast<BYTE>(b1 >> 4);
			m_lpResponseLevels[i].Level = static_cast<BYTE>(b1 & 0xf);
		}
	}
}

_Check_return_ wstring ConversationIndex::ToStringInternal()
{
	wstring szConversationIndex;

	wstring PropString;
	wstring AltPropString;
	FileTimeToString(&m_ftCurrent, PropString, AltPropString);
	auto szGUID = GUIDToString(&m_guid);
	szConversationIndex = formatmessage(IDS_CONVERSATIONINDEXHEADER,
		m_UnnamedByte,
		m_ftCurrent.dwLowDateTime,
		m_ftCurrent.dwHighDateTime,
		PropString.c_str(),
		szGUID.c_str());

	if (m_ulResponseLevels && m_lpResponseLevels)
	{
		for (ULONG i = 0; i < m_ulResponseLevels; i++)
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