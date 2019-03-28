#include <core/stdafx.h>
#include <core/smartview/PCL.h>
#include <core/interpret/guid.h>

namespace smartview
{
	SizedXID::SizedXID(const std::shared_ptr<binaryParser>& parser)
	{
		XidSize = blockT<BYTE>::parse(parser);
		NamespaceGuid = blockT<GUID>::parse(parser);
		auto cbLocalId = *XidSize - sizeof(GUID);
		if (parser->getSize() >= cbLocalId)
		{
			LocalID = blockBytes::parse(parser, cbLocalId);
		}
	}

	void PCL::parse()
	{
		auto cXID = 0;
		// Run through the parser once to count the number of flag structs
		// Must have at least 1 byte left to have another XID
		while (m_Parser->getSize() > sizeof(BYTE))
		{
			const auto XidSize = blockT<BYTE>::parse(m_Parser);
			if (m_Parser->getSize() >= *XidSize)
			{
				m_Parser->advance(*XidSize);
			}

			cXID++;
		}

		// Now we parse for real
		m_Parser->rewind();

		if (cXID && cXID < _MaxEntriesSmall)
		{
			m_lpXID.reserve(cXID);
			for (auto i = 0; i < cXID; i++)
			{
				auto oldSize = m_Parser->getSize();
				m_lpXID.emplace_back(std::make_shared<SizedXID>(m_Parser));
				auto newSize = m_Parser->getSize();
				if (newSize == 0 || newSize == oldSize) break;
			}
		}
	}

	void PCL::parseBlocks()
	{
		setRoot(L"Predecessor Change List:\r\n");
		addHeader(L"Count = %1!d!", m_lpXID.size());

		if (!m_lpXID.empty())
		{
			auto i = 0;
			for (const auto& xid : m_lpXID)
			{
				terminateBlock();
				addHeader(L"XID[%1!d!]:\r\n", i);
				addChild(xid->XidSize, L"XidSize = 0x%1!08X! = %1!d!\r\n", xid->XidSize->getData());
				addChild(
					xid->NamespaceGuid,
					L"NamespaceGuid = %1!ws!\r\n",
					guid::GUIDToString(xid->NamespaceGuid->getData()).c_str());
				addHeader(L"LocalId = ");
				addChild(xid->LocalID);

				i++;
			}
		}
	}
} // namespace smartview