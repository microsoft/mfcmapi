#include <core/stdafx.h>
#include <core/smartview/XID.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	void XID::parse()
	{
		m_NamespaceGuid = blockT<GUID>::parse(m_Parser);
		const auto cbLocalId = m_Parser->getSize();
		m_LocalID = blockBytes::parse(m_Parser, cbLocalId, cbLocalId);
	}

	void XID::parseBlocks()
	{
		setRoot(L"XID:\r\n");
		addChild(
			m_NamespaceGuid, L"NamespaceGuid = %1!ws!\r\n", guid::GUIDToString(m_NamespaceGuid->getData()).c_str());
		if (m_LocalID) addChild(m_LocalID, L"LocalId = %1!ws!", m_LocalID->toHexString(true).c_str());
	}
} // namespace smartview