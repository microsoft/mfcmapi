#include <core/stdafx.h>
#include <core/smartview/FlatEntryList.h>
#include <core/utility/strings.h>

namespace smartview
{
	void FlatEntryID::parse()
	{
		// Size here will be the length of the serialized entry ID
		// We'll have to round it up to a multiple of 4 to read off padding
		dwSize = blockT<DWORD>::parse(parser);
		const auto ulSize = min(*dwSize, parser->getSize());

		lpEntryID = block::parse<EntryIdStruct>(parser, ulSize, true);

		const auto dwPAD = 3 - (*dwSize + 3) % 4;
		if (dwPAD > 0)
		{
			padding = blockBytes::parse(parser, dwPAD);
		}
	}
	void FlatEntryID::parseBlocks()
	{
		if (dwSize->isSet())
		{
			addChild(dwSize, L"Size = 0x%1!08X!", dwSize->getData());
		}

		if (lpEntryID)
		{
			addChild(lpEntryID);
		}

		if (padding->isSet())
		{
			auto paddingHeader = create(L"Padding:");
			addChild(paddingHeader);
			paddingHeader->addChild(padding, padding->toHexString(true));
		}
	}

	void FlatEntryList::parse()
	{
		m_cEntries = blockT<DWORD>::parse(parser);

		// We read and report this, but ultimately, it's not used.
		m_cbEntries = blockT<DWORD>::parse(parser);

		if (*m_cEntries && *m_cEntries < _MaxEntriesLarge)
		{
			m_pEntryIDs.reserve(*m_cEntries);
			for (DWORD i = 0; i < *m_cEntries; i++)
			{
				m_pEntryIDs.emplace_back(block::parse<FlatEntryID>(parser, 0, false));
			}
		}
	}

	void FlatEntryList::parseBlocks()
	{
		setText(L"Flat Entry List");
		addChild(m_cEntries, L"cEntries = %1!d!", m_cEntries->getData());
		addChild(m_cbEntries, L"cbEntries = 0x%1!08X!", m_cbEntries->getData());

		auto i = DWORD{};
		for (const auto& entry : m_pEntryIDs)
		{
			if (entry->isSet())
			{
				addChild(entry, L"Entry[%1!d!]", i);
			}

			i++;
		}
	}
} // namespace smartview