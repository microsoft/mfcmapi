#include "stdafx.h"
#include "..\stdafx.h"
#include "RecipientRowStream.h"
#include "BinaryParser.h"
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

	CBinaryParser Parser(m_cbBin, m_lpBin);

	Parser.GetDWORD(&m_cVersion);
	Parser.GetDWORD(&m_cRowCount);

	if (m_cRowCount && m_cRowCount < _MaxEntriesSmall)
		m_lpAdrEntry = new ADRENTRY[m_cRowCount];

	if (m_lpAdrEntry)
	{
		memset(m_lpAdrEntry, 0, sizeof(ADRENTRY)*m_cRowCount);
		ULONG i = 0;

		for (i = 0; i < m_cRowCount; i++)
		{
			Parser.GetDWORD(&m_lpAdrEntry[i].cValues);
			Parser.GetDWORD(&m_lpAdrEntry[i].ulReserved1);

			if (m_lpAdrEntry[i].cValues && m_lpAdrEntry[i].cValues < _MaxEntriesSmall)
			{
				size_t cbOffset = Parser.GetCurrentOffset();
				size_t cbBytesRead = 0;
				m_lpAdrEntry[i].rgPropVals = BinToSPropValue(
					(ULONG)Parser.RemainingBytes(),
					m_lpBin + cbOffset,
					m_lpAdrEntry[i].cValues,
					&cbBytesRead,
					false);
				Parser.Advance(cbBytesRead);
			}
		}
	}

	m_JunkDataSize = Parser.GetRemainingData(&m_JunkData);
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