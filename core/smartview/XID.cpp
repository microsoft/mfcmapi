#include <core/stdafx.h>
#include <core/smartview/XID.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	void XID::parse()
	{
		m_NamespaceGuid = blockT<GUID>::parse(parser);
		const auto cbLocalId = parser->getSize();
		m_LocalID = blockBytes::parse(parser, cbLocalId, cbLocalId);
	}

	void XID::parseBlocks()
	{
		setText(L"XID");
		addChild(
			m_NamespaceGuid, L"NamespaceGuid = %1!ws!", guid::GUIDToString(m_NamespaceGuid->getData()).c_str());
		if (m_LocalID) addChild(m_LocalID, L"LocalId = %1!ws!", m_LocalID->toHexString(true).c_str());
	}
} // namespace smartview