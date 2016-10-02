#include "stdafx.h"
#include "PCL.h"
#include "String.h"
#include "InterpretProp2.h"

PCL::PCL(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_cXID = 0;
	m_lpXID = nullptr;
}

PCL::~PCL()
{
	if (m_cXID && m_lpXID)
	{
		for (DWORD i = 0; i < m_cXID; i++)
		{
			delete[] m_lpXID[i].LocalID;
		}

		delete[] m_lpXID;
	}
}

void PCL::Parse()
{
	m_cXID = 0;

	// Run through the parser once to count the number of flag structs
	for (;;)
	{
		// Must have at least 1 byte left to have another XID
		if (m_Parser.RemainingBytes() <= sizeof(BYTE)) break;

		BYTE XidSize = 0;
		m_Parser.GetBYTE(&XidSize);
		if (m_Parser.RemainingBytes() >= XidSize)
		{
			m_Parser.Advance(XidSize);
		}

		m_cXID++;
	}

	// Now we parse for real
	m_Parser.Rewind();

	if (m_cXID && m_cXID < _MaxEntriesSmall)
		m_lpXID = new SizedXID[m_cXID];

	if (m_lpXID)
	{
		memset(m_lpXID, 0, sizeof(SizedXID)*m_cXID);
		for (DWORD i = 0; i < m_cXID; i++)
		{
			m_Parser.GetBYTE(&m_lpXID[i].XidSize);
			m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), reinterpret_cast<LPBYTE>(&m_lpXID[i].NamespaceGuid));
			m_lpXID[i].cbLocalId = m_lpXID[i].XidSize - sizeof(GUID);
			if (m_Parser.RemainingBytes() < m_lpXID[i].cbLocalId) break;
			m_Parser.GetBYTES(m_lpXID[i].cbLocalId, m_lpXID[i].cbLocalId, &m_lpXID[i].LocalID);
		}
	}
}

_Check_return_ wstring PCL::ToStringInternal()
{
	auto szPCLString = formatmessage(IDS_PCLHEADER, m_cXID);

	if (m_cXID && m_lpXID)
	{
		ULONG i = 0;
		for (i = 0; i < m_cXID; i++)
		{
			SBinary sBin = { 0 };
			sBin.cb = static_cast<ULONG>(m_lpXID[i].cbLocalId);
			sBin.lpb = m_lpXID[i].LocalID;

			szPCLString += formatmessage(IDS_PCLXID,
				i,
				m_lpXID[i].XidSize,
				GUIDToString(&m_lpXID[i].NamespaceGuid).c_str(),
				BinToHexString(&sBin, true).c_str());
		}
	}

	return szPCLString;
}