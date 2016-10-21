#include "stdafx.h"
#include "XID.h"
#include "InterpretProp2.h"

XID::XID()
{
	m_NamespaceGuid = { 0 };
	m_cbLocalId = 0;
}

void XID::Parse()
{
	m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), (LPBYTE)&m_NamespaceGuid);
	m_cbLocalId = m_Parser.RemainingBytes();
	m_LocalID = m_Parser.GetBYTES(m_cbLocalId, m_cbLocalId);
}

_Check_return_ wstring XID::ToStringInternal()
{
	return formatmessage(IDS_XID,
		GUIDToString(&m_NamespaceGuid).c_str(),
		BinToHexString(m_LocalID, true).c_str());
}