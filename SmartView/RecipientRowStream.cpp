#include "stdafx.h"
#include "RecipientRowStream.h"
#include "String.h"
#include "PropertyStruct.h"

RecipientRowStream::RecipientRowStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	Init(cbBin, lpBin);
	m_cVersion = 0;
	m_cRowCount = 0;
	m_lpAdrEntry = nullptr;
}

RecipientRowStream::~RecipientRowStream()
{
	if (m_lpAdrEntry && m_cRowCount)
	{
		for (DWORD i = 0; i < m_cRowCount; i++)
		{
			DeleteSPropVal(m_lpAdrEntry[i].cValues, m_lpAdrEntry[i].rgPropVals);
		}
	}

	delete[] m_lpAdrEntry;
}

void RecipientRowStream::Parse()
{
	m_Parser.GetDWORD(&m_cVersion);
	m_Parser.GetDWORD(&m_cRowCount);

	if (m_cRowCount && m_cRowCount < _MaxEntriesSmall)
		m_lpAdrEntry = new ADRENTRY[m_cRowCount];

	if (m_lpAdrEntry)
	{
		memset(m_lpAdrEntry, 0, sizeof(ADRENTRY)*m_cRowCount);

		for (DWORD i = 0; i < m_cRowCount; i++)
		{
			m_Parser.GetDWORD(&m_lpAdrEntry[i].cValues);
			m_Parser.GetDWORD(&m_lpAdrEntry[i].ulReserved1);

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

	szRecipientRowStream = formatmessage(
		IDS_RECIPIENTROWSTREAMHEADER,
		m_cVersion,
		m_cRowCount);
	if (m_lpAdrEntry && m_cRowCount)
	{
		for (DWORD i = 0; i < m_cRowCount; i++)
		{
			szRecipientRowStream += formatmessage(
				IDS_RECIPIENTROWSTREAMROW,
				i,
				m_lpAdrEntry[i].cValues,
				m_lpAdrEntry[i].ulReserved1);

			szRecipientRowStream += PropsToString(m_lpAdrEntry[i].cValues, m_lpAdrEntry[i].rgPropVals);
		}
	}

	return szRecipientRowStream;
}