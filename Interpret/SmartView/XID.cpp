#include <StdAfx.h>
#include <Interpret/SmartView/XID.h>
#include <Interpret/Guids.h>

namespace smartview
{
	void XID::Parse()
	{
		m_NamespaceGuid = m_Parser.Get<GUID>();
		m_cbLocalId = m_Parser.RemainingBytes();
		m_LocalID = m_Parser.GetBYTES(m_cbLocalId, m_cbLocalId);
	}

	void XID::ParseBlocks()
	{
		addHeader(L"XID:\r\n");
		addBlock(m_NamespaceGuid, L"NamespaceGuid = %1!ws!\r\n", guid::GUIDToString(m_NamespaceGuid.getData()).c_str());
		addBlock(m_LocalID, L"LocalId = %1!ws!", strings::BinToHexString(m_LocalID, true).c_str());
	}
} // namespace smartview