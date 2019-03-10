#include <core/stdafx.h>
#include <core/smartview/FlatEntryList.h>
#include <core/utility/strings.h>

namespace smartview
{
	FlatEntryID::FlatEntryID(std::shared_ptr<binaryParser> parser)
	{
		// Size here will be the length of the serialized entry ID
		// We'll have to round it up to a multiple of 4 to read off padding
		dwSize = parser->Get<DWORD>();
		const auto ulSize = min(dwSize, parser->RemainingBytes());

		lpEntryID.parse(parser, ulSize, true);

		const auto dwPAD = 3 - (dwSize + 3) % 4;
		if (dwPAD > 0)
		{
			padding.parse(parser, dwPAD);
		}
	}

	void FlatEntryList::Parse()
	{
		m_cEntries = m_Parser->Get<DWORD>();

		// We read and report this, but ultimately, it's not used.
		m_cbEntries = m_Parser->Get<DWORD>();

		if (m_cEntries && m_cEntries < _MaxEntriesLarge)
		{
			m_pEntryIDs.reserve(m_cEntries);
			for (DWORD iFlatEntryList = 0; iFlatEntryList < m_cEntries; iFlatEntryList++)
			{
				m_pEntryIDs.emplace_back(std::make_shared<FlatEntryID>(m_Parser));
			}
		}
	}

	void FlatEntryList::ParseBlocks()
	{
		setRoot(L"Flat Entry List\r\n");
		addBlock(m_cEntries, L"cEntries = %1!d!\r\n", m_cEntries.getData());
		addBlock(m_cbEntries, L"cbEntries = 0x%1!08X!", m_cbEntries.getData());

		auto iFlatEntryList = DWORD{};
		for (const auto& entry : m_pEntryIDs)
		{
			terminateBlock();
			addBlankLine();
			addHeader(L"Entry[%1!d!] ", iFlatEntryList);
			addBlock(entry->dwSize, L"Size = 0x%1!08X!", entry->dwSize.getData());

			if (entry->lpEntryID.hasData())
			{
				terminateBlock();
				addBlock(entry->lpEntryID.getBlock());
			}

			if (!entry->padding.empty())
			{
				terminateBlock();
				addHeader(L"Entry[%1!d!] Padding:\r\n", iFlatEntryList);
				addBlock(entry->padding, strings::BinToHexString(entry->padding, true));
			}

			iFlatEntryList++;
		}
	}
} // namespace smartview