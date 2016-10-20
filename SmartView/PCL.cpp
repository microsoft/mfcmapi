#include "stdafx.h"
#include "PCL.h"
#include "InterpretProp2.h"

PCL::PCL()
{
	m_cXID = 0;
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
	{
		for (DWORD i = 0; i < m_cXID; i++)
		{
			SizedXID sizedXID;
			m_Parser.GetBYTE(&sizedXID.XidSize);
			m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), reinterpret_cast<LPBYTE>(&sizedXID.NamespaceGuid));
			sizedXID.cbLocalId = sizedXID.XidSize - sizeof(GUID);
			if (m_Parser.RemainingBytes() < sizedXID.cbLocalId) break;
			sizedXID.LocalID = m_Parser.GetBYTES(sizedXID.cbLocalId);
			m_lpXID.push_back(sizedXID);
		}
	}
}

_Check_return_ wstring PCL::ToStringInternal()
{
	auto szPCLString = formatmessage(IDS_PCLHEADER, m_cXID);

	if (m_lpXID.size())
	{
		for (ULONG i = 0; i < m_lpXID.size(); i++)
		{
			szPCLString += formatmessage(IDS_PCLXID,
				i,
				m_lpXID[i].XidSize,
				GUIDToString(&m_lpXID[i].NamespaceGuid).c_str(),
				BinToHexString(m_lpXID[i].LocalID, true).c_str());
		}
	}

	return szPCLString;
}