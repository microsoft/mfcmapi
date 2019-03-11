#include <core/stdafx.h>
#include <core/smartview/XID.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	void XID::Parse()
	{
		m_NamespaceGuid.parse<GUID>(m_Parser);
		m_cbLocalId = m_Parser->RemainingBytes();
		m_LocalID.parse(m_Parser, m_cbLocalId, m_cbLocalId);
	}

	void XID::ParseBlocks()
	{
		setRoot(L"XID:\r\n");
		addBlock(m_NamespaceGuid, L"NamespaceGuid = %1!ws!\r\n", guid::GUIDToString(m_NamespaceGuid.getData()).c_str());
		addBlock(m_LocalID, L"LocalId = %1!ws!", strings::BinToHexString(m_LocalID, true).c_str());
	}
} // namespace smartview