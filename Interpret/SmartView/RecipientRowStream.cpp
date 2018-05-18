#include "stdafx.h"
#include "RecipientRowStream.h"
#include <Interpret/String.h>
#include "PropertyStruct.h"

RecipientRowStream::RecipientRowStream()
{
	m_cVersion = 0;
	m_cRowCount = 0;
	m_lpAdrEntry = nullptr;
}

void RecipientRowStream::Parse()
{
	m_cVersion = m_Parser.Get<DWORD>();
	m_cRowCount = m_Parser.Get<DWORD>();

	if (m_cRowCount && m_cRowCount < _MaxEntriesSmall)
		m_lpAdrEntry = reinterpret_cast<ADRENTRY*>(AllocateArray(m_cRowCount, sizeof ADRENTRY));

	if (m_lpAdrEntry)
	{
		for (DWORD i = 0; i < m_cRowCount; i++)
		{
			m_lpAdrEntry[i].cValues = m_Parser.Get<DWORD>();
			m_lpAdrEntry[i].ulReserved1 = m_Parser.Get<DWORD>();

			if (m_lpAdrEntry[i].cValues && m_lpAdrEntry[i].cValues < _MaxEntriesSmall)
			{
				m_lpAdrEntry[i].rgPropVals = BinToSPropValue(
					m_lpAdrEntry[i].cValues,
					false);
			}
		}
	}
}

_Check_return_ wstring RecipientRowStream::ToStringInternal()
{
	wstring szRecipientRowStream;

	szRecipientRowStream = strings::formatmessage(
		IDS_RECIPIENTROWSTREAMHEADER,
		m_cVersion,
		m_cRowCount);
	if (m_lpAdrEntry && m_cRowCount)
	{
		for (DWORD i = 0; i < m_cRowCount; i++)
		{
			szRecipientRowStream += strings::formatmessage(
				IDS_RECIPIENTROWSTREAMROW,
				i,
				m_lpAdrEntry[i].cValues,
				m_lpAdrEntry[i].ulReserved1);

			szRecipientRowStream += PropsToString(m_lpAdrEntry[i].cValues, m_lpAdrEntry[i].rgPropVals);
		}
	}

	return szRecipientRowStream;
}