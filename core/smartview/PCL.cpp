#include <core/stdafx.h>
#include <core/smartview/PCL.h>
#include <core/interpret/guid.h>

namespace smartview
{
	void PCL::Parse()
	{
		// Run through the parser once to count the number of flag structs
		for (;;)
		{
			// Must have at least 1 byte left to have another XID
			if (m_Parser.RemainingBytes() <= sizeof(BYTE)) break;

			const auto XidSize = m_Parser.Get<BYTE>();
			if (m_Parser.RemainingBytes() >= XidSize)
			{
				m_Parser.advance(XidSize);
			}

			m_cXID++;
		}

		// Now we parse for real
		m_Parser.rewind();

		if (m_cXID && m_cXID < _MaxEntriesSmall)
		{
			m_lpXID.reserve(m_cXID);
			for (DWORD i = 0; i < m_cXID; i++)
			{
				SizedXID sizedXID;
				sizedXID.XidSize = m_Parser.Get<BYTE>();
				sizedXID.NamespaceGuid = m_Parser.Get<GUID>();
				sizedXID.cbLocalId = sizedXID.XidSize - sizeof(GUID);
				if (m_Parser.RemainingBytes() < sizedXID.cbLocalId)
				{
					m_lpXID.push_back(sizedXID);
					break;
				}

				sizedXID.LocalID = m_Parser.GetBYTES(sizedXID.cbLocalId);
				m_lpXID.push_back(sizedXID);
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
			for (auto& xid : m_lpXID)
			{
				terminateBlock();
				addHeader(L"XID[%1!d!]:\r\n", i++);
				addBlock(xid.XidSize, L"XidSize = 0x%1!08X! = %1!d!\r\n", xid.XidSize.getData());
				addBlock(
					xid.NamespaceGuid,
					L"NamespaceGuid = %1!ws!\r\n",
					guid::GUIDToString(xid.NamespaceGuid.getData()).c_str());
				addHeader(L"LocalId = ");
				addBlock(xid.LocalID);
			}
		}
	}
} // namespace smartview