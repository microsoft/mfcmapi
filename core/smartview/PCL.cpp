#include <core/stdafx.h>
#include <core/smartview/PCL.h>
#include <core/interpret/guid.h>

namespace smartview
{
	SizedXID::SizedXID(std::shared_ptr<binaryParser>& parser)
	{
		XidSize.parse<BYTE>(parser);
		NamespaceGuid.parse<GUID>(parser);
		cbLocalId = XidSize - sizeof(GUID);
		if (parser->RemainingBytes() >= cbLocalId)
		{
			LocalID.parse(parser, cbLocalId);
		}
	}

	void PCL::Parse()
	{
		// Run through the parser once to count the number of flag structs
		// Must have at least 1 byte left to have another XID
		while (m_Parser->RemainingBytes() > sizeof(BYTE))
		{
			const auto& XidSize = blockT<BYTE>(m_Parser);
			if (m_Parser->RemainingBytes() >= XidSize)
			{
				m_Parser->advance(XidSize);
			}

			m_cXID++;
		}

		// Now we parse for real
		m_Parser->rewind();

		if (m_cXID && m_cXID < _MaxEntriesSmall)
		{
			m_lpXID.reserve(m_cXID);
			for (DWORD i = 0; i < m_cXID; i++)
			{
				m_lpXID.emplace_back(std::make_shared<SizedXID>(m_Parser));
			}
		}
	}

	void PCL::ParseBlocks()
	{
		setRoot(L"Predecessor Change List:\r\n");
		addHeader(L"Count = %1!d!", m_cXID);

		if (!m_lpXID.empty())
		{
			auto i = 0;
			for (const auto& xid : m_lpXID)
			{
				terminateBlock();
				addHeader(L"XID[%1!d!]:\r\n", i++);
				addChild(xid->XidSize, L"XidSize = 0x%1!08X! = %1!d!\r\n", xid->XidSize.getData());
				addChild(
					xid->NamespaceGuid,
					L"NamespaceGuid = %1!ws!\r\n",
					guid::GUIDToString(xid->NamespaceGuid.getData()).c_str());
				addHeader(L"LocalId = ");
				addChild(xid->LocalID);
			}
		}
	}
} // namespace smartview