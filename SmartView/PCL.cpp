#include "stdafx.h"
#include "..\stdafx.h"
#include "PCL.h"
#include "BinaryParser.h"
#include "..\String.h"
#include "..\ParseProperty.h"

PCL::PCL(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_cXID = 0;
	m_lpXID = NULL;
}

PCL::~PCL()
{
	if (m_cXID && m_lpXID)
	{
		ULONG i = 0;

		for (i = 0; i < m_cXID; i++)
		{
			delete[] m_lpXID[i].LocalID;
		}

		delete[] m_lpXID;
	}
}

void PCL::Parse()
{
	if (!m_lpBin) return;

	m_cXID = 0;

	// Run through the parser once to count the number of flag structs
	CBinaryParser Parser2(m_cbBin, m_lpBin);
	for (;;)
	{
		// Must have at least 1 byte left to have another XID
		if (Parser2.RemainingBytes() <= sizeof(BYTE)) break;

		BYTE XidSize = 0;
		Parser2.GetBYTE(&XidSize);
		if (Parser2.RemainingBytes() >= XidSize)
		{
			Parser2.Advance(XidSize);
		}

		m_cXID++;
	}

	if (m_cXID && m_cXID < _MaxEntriesSmall)
		m_lpXID = new SizedXID[m_cXID];

	if (m_lpXID)
	{
		memset(m_lpXID, 0, sizeof(SizedXID)*m_cXID);
		ULONG i = 0;

		for (i = 0; i < m_cXID; i++)
		{
			m_Parser.GetBYTE(&m_lpXID[i].XidSize);
			m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), (LPBYTE)&m_lpXID[i].NamespaceGuid);
			m_lpXID[i].cbLocalId = m_lpXID[i].XidSize - sizeof(GUID);
			if (m_Parser.RemainingBytes() < m_lpXID[i].cbLocalId) break;
			m_Parser.GetBYTES(m_lpXID[i].cbLocalId, m_lpXID[i].cbLocalId, &m_lpXID[i].LocalID);
		}
	}
}

_Check_return_ LPWSTR PCL::ToString()
{
	Parse();

	wstring szPCLString;
	wstring szTmp;

	szPCLString = formatmessage(IDS_PCLHEADER, m_cXID);

	if (m_cXID && m_lpXID)
	{
		ULONG i = 0;
		for (i = 0; i < m_cXID; i++)
		{
			wstring szGUID = GUIDToWstring(&m_lpXID[i].NamespaceGuid);

			SBinary sBin = { 0 };
			sBin.cb = (ULONG)m_lpXID[i].cbLocalId;
			sBin.lpb = m_lpXID[i].LocalID;

			szTmp = formatmessage(IDS_PCLXID,
				i,
				m_lpXID[i].XidSize,
				szGUID.c_str(),
				BinToHexString(&sBin, true).c_str());
			szPCLString += szTmp;
		}
	}

	szPCLString += JunkDataToString();

	return wstringToLPWSTR(szPCLString);
}