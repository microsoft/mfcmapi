#include <core/stdafx.h>
#include <core/smartview/FlatEntryList.h>
#include <core/utility/strings.h>

namespace smartview
{
	FlatEntryID::FlatEntryID(const std::shared_ptr<binaryParser>& parser)
	{
		// Size here will be the length of the serialized entry ID
		// We'll have to round it up to a multiple of 4 to read off padding
		dwSize = blockT<DWORD>::parse(parser);
		const auto ulSize = min(*dwSize, parser->RemainingBytes());

		lpEntryID.parse(parser, ulSize, true);

		const auto dwPAD = 3 - (*dwSize + 3) % 4;
		if (dwPAD > 0)
		{
			padding = blockBytes::parse(parser, dwPAD);
		}
	}

	void FlatEntryList::Parse()
	{
		m_cEntries = blockT<DWORD>::parse(m_Parser);

		// We read and report this, but ultimately, it's not used.
		m_cbEntries = blockT<DWORD>::parse(m_Parser);

		if (*m_cEntries && *m_cEntries < _MaxEntriesLarge)
		{
			m_pEntryIDs.reserve(*m_cEntries);
			for (DWORD i = 0; i < *m_cEntries; i++)
			{
				m_pEntryIDs.emplace_back(std::make_shared<FlatEntryID>(m_Parser));
			}
		}
	}

	void FlatEntryList::ParseBlocks()
	{
		setRoot(L"Flat Entry List\r\n");
		addChild(m_cEntries, L"cEntries = %1!d!\r\n", m_cEntries->getData());
		addChild(m_cbEntries, L"cbEntries = 0x%1!08X!", m_cbEntries->getData());

		auto i = DWORD{};
		for (const auto& entry : m_pEntryIDs)
		{
			terminateBlock();
			addBlankLine();
			addHeader(L"Entry[%1!d!] ", i);
			addChild(entry->dwSize, L"Size = 0x%1!08X!", entry->dwSize->getData());

			if (entry->lpEntryID.hasData())
			{
				terminateBlock();
				addChild(entry->lpEntryID.getBlock());
			}

			if (!entry->padding->empty())
			{
				terminateBlock();
				addHeader(L"Entry[%1!d!] Padding:\r\n", i);
				if (entry->padding) addChild(entry->padding, strings::BinToHexString(*entry->padding, true));
			}

			i++;
		}
	}
} // namespace smartview