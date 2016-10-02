#include "stdafx.h"
#include "XID.h"
#include "InterpretProp2.h"

XID::XID(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_NamespaceGuid = { 0 };
	m_cbLocalId = 0;
	m_LocalID = nullptr;
}

XID::~XID()
{
	if (m_LocalID) delete[] m_LocalID;
}

void XID::Parse()
{
	m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), (LPBYTE)&m_NamespaceGuid);
	m_cbLocalId = m_Parser.RemainingBytes();
	m_Parser.GetBYTES(m_cbLocalId, m_cbLocalId, &m_LocalID);
}

_Check_return_ wstring XID::ToStringInternal()
{
	SBinary sBin = { 0 };
	sBin.cb = (ULONG)m_cbLocalId;
	sBin.lpb = m_LocalID;

	return formatmessage(IDS_XID,
		GUIDToString(&m_NamespaceGuid).c_str(),
		BinToHexString(&sBin, true).c_str());
}