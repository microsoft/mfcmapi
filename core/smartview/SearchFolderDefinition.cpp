#include <core/stdafx.h>
#include <core/smartview/SearchFolderDefinition.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>

namespace smartview
{
	AddressListEntryStruct::AddressListEntryStruct(std::shared_ptr<binaryParser> parser)
	{
		PropertyCount.parse<DWORD>(parser);
		Pad.parse<DWORD>(parser);
		if (PropertyCount)
		{
			Props.SetMaxEntries(PropertyCount);
			Props.parse(parser, false);
		}
	}

	void SearchFolderDefinition::Parse()
	{
		m_Version.parse<DWORD>(m_Parser);
		m_Flags.parse<DWORD>(m_Parser);
		m_NumericSearch.parse<DWORD>(m_Parser);

		m_TextSearchLength.parse<BYTE>(m_Parser);
		size_t cchTextSearch = m_TextSearchLength;
		if (255 == m_TextSearchLength)
		{
			m_TextSearchLengthExtended.parse<WORD>(m_Parser);
			cchTextSearch = m_TextSearchLengthExtended;
		}

		if (cchTextSearch)
		{
			m_TextSearch.parse(m_Parser, cchTextSearch);
		}

		m_SkipLen1.parse<DWORD>(m_Parser);
		m_SkipBytes1.parse(m_Parser, m_SkipLen1, _MaxBytes);

		m_DeepSearch.parse<DWORD>(m_Parser);

		m_FolderList1Length.parse<BYTE>(m_Parser);
		size_t cchFolderList1 = m_FolderList1Length;
		if (255 == m_FolderList1Length)
		{
			m_FolderList1LengthExtended.parse<WORD>(m_Parser);
			cchFolderList1 = m_FolderList1LengthExtended;
		}

		if (cchFolderList1)
		{
			m_FolderList1.parse(m_Parser, cchFolderList1);
		}

		m_FolderList2Length.parse<DWORD>(m_Parser);

		if (m_FolderList2Length)
		{
			m_FolderList2.parse(m_Parser, m_FolderList2Length, true);
		}

		if (SFST_BINARY & m_Flags)
		{
			m_AddressCount.parse<DWORD>(m_Parser);
			if (m_AddressCount)
			{
				if (m_AddressCount < _MaxEntriesSmall)
				{
					m_Addresses.reserve(m_AddressCount);
					for (DWORD i = 0; i < m_AddressCount; i++)
					{
						m_Addresses.emplace_back(std::make_shared<AddressListEntryStruct>(m_Parser));
					}
				}
			}
		}

		m_SkipLen2.parse<DWORD>(m_Parser);
		m_SkipBytes2.parse(m_Parser, m_SkipLen2, _MaxBytes);

		if (SFST_MRES & m_Flags)
		{
			m_Restriction = std::make_shared<RestrictionStruct>(false, true);
			m_Restriction->SmartViewParser::parse(m_Parser, false);
		}

		if (SFST_FILTERSTREAM & m_Flags)
		{
			const auto cbRemainingBytes = m_Parser->RemainingBytes();
			// Since the format for SFST_FILTERSTREAM isn't documented, just assume that everything remaining
			// is part of this bucket. We leave DWORD space for the final skip block, which should be empty
			if (cbRemainingBytes > sizeof DWORD)
			{
				m_AdvancedSearchBytes.parse(m_Parser, cbRemainingBytes - sizeof DWORD);
			}
		}

		m_SkipLen3.parse<DWORD>(m_Parser);
		if (m_SkipLen3)
		{
			m_SkipBytes3.parse(m_Parser, m_SkipLen3, _MaxBytes);
		}
	}

	void SearchFolderDefinition::ParseBlocks()
	{
		setRoot(L"Search Folder Definition:\r\n");
		addBlock(m_Version, L"Version = 0x%1!08X!\r\n", m_Version.getData());
		addBlock(
			m_Flags,
			L"Flags = 0x%1!08X! = %2!ws!\r\n",
			m_Flags.getData(),
			InterpretNumberAsStringProp(m_Flags, PR_WB_SF_STORAGE_TYPE).c_str());
		addBlock(m_NumericSearch, L"Numeric Search = 0x%1!08X!\r\n", m_NumericSearch.getData());
		addBlock(m_TextSearchLength, L"Text Search Length = 0x%1!02X!", m_TextSearchLength.getData());

		if (m_TextSearchLength)
		{
			terminateBlock();
			addBlock(
				m_TextSearchLengthExtended,
				L"Text Search Length Extended = 0x%1!04X!\r\n",
				m_TextSearchLengthExtended.getData());
			addHeader(L"Text Search = ");
			if (!m_TextSearch.empty())
			{
				addBlock(m_TextSearch, m_TextSearch.c_str());
			}
		}

		terminateBlock();
		addBlock(m_SkipLen1, L"SkipLen1 = 0x%1!08X!", m_SkipLen1.getData());

		if (m_SkipLen1)
		{
			terminateBlock();
			addHeader(L"SkipBytes1 = ");

			addBlock(m_SkipBytes1);
		}

		terminateBlock();
		addBlock(m_DeepSearch, L"Deep Search = 0x%1!08X!\r\n", m_DeepSearch.getData());
		addBlock(m_FolderList1Length, L"Folder List 1 Length = 0x%1!02X!", m_FolderList1Length.getData());

		if (m_FolderList1Length)
		{
			terminateBlock();
			addBlock(
				m_FolderList1LengthExtended,
				L"Folder List 1 Length Extended = 0x%1!04X!\r\n",
				m_FolderList1LengthExtended.getData());
			addHeader(L"Folder List 1 = ");

			if (!m_FolderList1.empty())
			{
				addBlock(m_FolderList1, m_FolderList1.c_str());
			}
		}

		terminateBlock();
		addBlock(m_FolderList2Length, L"Folder List 2 Length = 0x%1!08X!", m_FolderList2Length.getData());

		if (m_FolderList2Length)
		{
			terminateBlock();
			addHeader(L"FolderList2 = \r\n");
			addBlock(m_FolderList2.getBlock());
		}

		if (SFST_BINARY & m_Flags)
		{
			terminateBlock();
			addBlock(m_AddressCount, L"AddressCount = 0x%1!08X!", m_AddressCount.getData());

			auto i = DWORD{};
			for (const auto& address : m_Addresses)
			{
				terminateBlock();
				addBlock(
					address->PropertyCount,
					L"Addresses[%1!d!].PropertyCount = 0x%2!08X!\r\n",
					i,
					address->PropertyCount.getData());
				addBlock(address->Pad, L"Addresses[%1!d!].Pad = 0x%2!08X!\r\n", i, address->Pad.getData());

				addHeader(L"Properties[%1!d!]:\r\n", i);
				addBlock(address->Props.getBlock());
				i++;
			}
		}

		terminateBlock();
		addBlock(m_SkipLen2, L"SkipLen2 = 0x%1!08X!", m_SkipLen2.getData());

		if (m_SkipLen2)
		{
			terminateBlock();
			addHeader(L"SkipBytes2 = ");
			addBlock(m_SkipBytes2);
		}

		if (m_Restriction && m_Restriction->hasData())
		{
			terminateBlock();
			addBlock(m_Restriction->getBlock());
		}

		if (SFST_FILTERSTREAM & m_Flags)
		{
			terminateBlock();
			addBlock(m_AdvancedSearchBytes, L"AdvancedSearchLen = 0x%1!08X!", m_AdvancedSearchBytes.size());

			if (!m_AdvancedSearchBytes.empty())
			{
				terminateBlock();
				addHeader(L"AdvancedSearchBytes = ");
				addBlock(m_AdvancedSearchBytes);
			}
		}

		terminateBlock();
		addBlock(m_SkipLen3, L"SkipLen3 = 0x%1!08X!", m_SkipLen3.getData());

		if (m_SkipLen3)
		{
			terminateBlock();
			addHeader(L"SkipBytes3 = ");
			addBlock(m_SkipBytes3);
		}
	}
} // namespace smartview