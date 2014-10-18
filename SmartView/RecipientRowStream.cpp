#include "stdafx.h"
#include "..\stdafx.h"
#include "RecipientRowStream.h"
#include "..\String.h"
#include "SmartView.h"

RecipientRowStream::RecipientRowStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_cVersion = 0;
	m_cRowCount = 0;
	m_lpAdrEntry = 0;
}

RecipientRowStream::~RecipientRowStream()
{
	if (m_lpAdrEntry && m_cRowCount)
	{
		ULONG i = 0;

		for (i = 0; i < m_cRowCount; i++)
		{
			DeleteSPropVal(m_lpAdrEntry[i].cValues, m_lpAdrEntry[i].rgPropVals);
		}
	}

	delete[] m_lpAdrEntry;
}

void RecipientRowStream::Parse()
{
	if (!m_lpBin) return;

	m_Parser.GetDWORD(&m_cVersion);
	m_Parser.GetDWORD(&m_cRowCount);

	if (m_cRowCount && m_cRowCount < _MaxEntriesSmall)
		m_lpAdrEntry = new ADRENTRY[m_cRowCount];

	if (m_lpAdrEntry)
	{
		memset(m_lpAdrEntry, 0, sizeof(ADRENTRY)*m_cRowCount);
		ULONG i = 0;

		for (i = 0; i < m_cRowCount; i++)
		{
			m_Parser.GetDWORD(&m_lpAdrEntry[i].cValues);
			m_Parser.GetDWORD(&m_lpAdrEntry[i].ulReserved1);

			if (m_lpAdrEntry[i].cValues && m_lpAdrEntry[i].cValues < _MaxEntriesSmall)
			{
				size_t cbOffset = m_Parser.GetCurrentOffset();
				size_t cbBytesRead = 0;
				m_lpAdrEntry[i].rgPropVals = BinToSPropValue(
					(ULONG)m_Parser.RemainingBytes(),
					m_lpBin + cbOffset,
					m_lpAdrEntry[i].cValues,
					&cbBytesRead,
					false);
				m_Parser.Advance(cbBytesRead);
			}
		}
	}
}

_Check_return_ LPWSTR RecipientRowStream::ToString()
{
	Parse();

	wstring szRecipientRowStream;
	wstring szTmp;

	szRecipientRowStream= formatmessage(
		IDS_RECIPIENTROWSTREAMHEADER,
		m_cVersion,
		m_cRowCount);
	if (m_lpAdrEntry && m_cRowCount)
	{
		ULONG i = 0;
		for (i = 0; i < m_cRowCount; i++)
		{
			szTmp= formatmessage(
				IDS_RECIPIENTROWSTREAMROW,
				i,
				m_lpAdrEntry[i].cValues,
				m_lpAdrEntry[i].ulReserved1);
			szRecipientRowStream += szTmp;

			PropertyStruct psPropStruct = { 0 };
			psPropStruct.PropCount = m_lpAdrEntry[i].cValues;
			psPropStruct.Prop = m_lpAdrEntry[i].rgPropVals;

			LPWSTR szProps = PropertyStructToString(&psPropStruct);
			szRecipientRowStream += szProps;
			delete[] szProps;
		}
	}

	szRecipientRowStream += JunkDataToString();

	return wstringToLPWSTR(szRecipientRowStream);
}