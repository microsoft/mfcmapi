#include <core/stdafx.h>
#include <core/smartview/SearchFolderDefinition.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>

namespace smartview
{
	AddressListEntryStruct::AddressListEntryStruct(const std::shared_ptr<binaryParser>& parser)
	{
		PropertyCount = blockT<DWORD>::parse(parser);
		Pad = blockT<DWORD>::parse(parser);
		if (*PropertyCount)
		{
			Props.SetMaxEntries(*PropertyCount);
			Props.SmartViewParser::parse(parser, false);
		}
	}

	void SearchFolderDefinition::Parse()
	{
		m_Version = blockT<DWORD>::parse(m_Parser);
		m_Flags = blockT<DWORD>::parse(m_Parser);
		m_NumericSearch = blockT<DWORD>::parse(m_Parser);

		m_TextSearchLength = blockT<BYTE>::parse(m_Parser);
		size_t cchTextSearch = *m_TextSearchLength;
		if (*m_TextSearchLength == 255)
		{
			m_TextSearchLengthExtended = blockT<WORD>::parse(m_Parser);
			cchTextSearch = *m_TextSearchLengthExtended;
		}

		if (cchTextSearch)
		{
			m_TextSearch = blockStringW::parse(m_Parser, cchTextSearch);
		}

		m_SkipLen1 = blockT<DWORD>::parse(m_Parser);
		m_SkipBytes1 = blockBytes::parse(m_Parser, *m_SkipLen1, _MaxBytes);

		m_DeepSearch = blockT<DWORD>::parse(m_Parser);

		m_FolderList1Length = blockT<BYTE>::parse(m_Parser);
		size_t cchFolderList1 = *m_FolderList1Length;
		if (*m_FolderList1Length == 255)
		{
			m_FolderList1LengthExtended = blockT<WORD>::parse(m_Parser);
			cchFolderList1 = *m_FolderList1LengthExtended;
		}

		if (cchFolderList1)
		{
			m_FolderList1 = blockStringW::parse(m_Parser, cchFolderList1);
		}

		m_FolderList2Length = blockT<DWORD>::parse(m_Parser);

		if (m_FolderList2Length)
		{
			m_FolderList2.parse(m_Parser, *m_FolderList2Length, true);
		}

		if (*m_Flags & SFST_BINARY)
		{
			m_AddressCount = blockT<DWORD>::parse(m_Parser);
			if (*m_AddressCount)
			{
				if (*m_AddressCount < _MaxEntriesSmall)
				{
					m_Addresses.reserve(*m_AddressCount);
					for (DWORD i = 0; i < *m_AddressCount; i++)
					{
						m_Addresses.emplace_back(std::make_shared<AddressListEntryStruct>(m_Parser));
					}
				}
			}
		}

		m_SkipLen2 = blockT<DWORD>::parse(m_Parser);
		m_SkipBytes2 = blockBytes::parse(m_Parser, *m_SkipLen2, _MaxBytes);

		if (*m_Flags & SFST_MRES)
		{
			m_Restriction = std::make_shared<RestrictionStruct>(false, true);
			m_Restriction->SmartViewParser::parse(m_Parser, false);
		}

		if (*m_Flags & SFST_FILTERSTREAM)
		{
			const auto cbRemainingBytes = m_Parser->getSize();
			// Since the format for SFST_FILTERSTREAM isn't documented, just assume that everything remaining
			// is part of this bucket. We leave DWORD space for the final skip block, which should be empty
			if (cbRemainingBytes > sizeof DWORD)
			{
				m_AdvancedSearchBytes = blockBytes::parse(m_Parser, cbRemainingBytes - sizeof DWORD);
			}
		}

		m_SkipLen3 = blockT<DWORD>::parse(m_Parser);
		if (m_SkipLen3)
		{
			m_SkipBytes3 = blockBytes::parse(m_Parser, *m_SkipLen3, _MaxBytes);
		}
	}

	void SearchFolderDefinition::ParseBlocks()
	{
		setRoot(L"Search Folder Definition:\r\n");
		addChild(m_Version, L"Version = 0x%1!08X!\r\n", m_Version->getData());
		addChild(
			m_Flags,
			L"Flags = 0x%1!08X! = %2!ws!\r\n",
			m_Flags->getData(),
			InterpretNumberAsStringProp(*m_Flags, PR_WB_SF_STORAGE_TYPE).c_str());
		addChild(m_NumericSearch, L"Numeric Search = 0x%1!08X!\r\n", m_NumericSearch->getData());
		addChild(m_TextSearchLength, L"Text Search Length = 0x%1!02X!", m_TextSearchLength->getData());

		if (*m_TextSearchLength)
		{
			terminateBlock();
			addChild(
				m_TextSearchLengthExtended,
				L"Text Search Length Extended = 0x%1!04X!\r\n",
				m_TextSearchLengthExtended->getData());
			addHeader(L"Text Search = ");
			if (!m_TextSearch->empty())
			{
				addChild(m_TextSearch, m_TextSearch->c_str());
			}
		}

		terminateBlock();
		addChild(m_SkipLen1, L"SkipLen1 = 0x%1!08X!", m_SkipLen1->getData());

		if (*m_SkipLen1)
		{
			terminateBlock();
			addHeader(L"SkipBytes1 = ");

			addChild(m_SkipBytes1);
		}

		terminateBlock();
		addChild(m_DeepSearch, L"Deep Search = 0x%1!08X!\r\n", m_DeepSearch->getData());
		addChild(m_FolderList1Length, L"Folder List 1 Length = 0x%1!02X!", m_FolderList1Length->getData());

		if (*m_FolderList1Length)
		{
			terminateBlock();
			addChild(
				m_FolderList1LengthExtended,
				L"Folder List 1 Length Extended = 0x%1!04X!\r\n",
				m_FolderList1LengthExtended->getData());
			addHeader(L"Folder List 1 = ");

			if (!m_FolderList1->empty())
			{
				addChild(m_FolderList1, m_FolderList1->c_str());
			}
		}

		terminateBlock();
		addChild(m_FolderList2Length, L"Folder List 2 Length = 0x%1!08X!", m_FolderList2Length->getData());

		if (*m_FolderList2Length)
		{
			terminateBlock();
			addHeader(L"FolderList2 = \r\n");
			addChild(m_FolderList2.getBlock());
		}

		if (*m_Flags & SFST_BINARY)
		{
			terminateBlock();
			addChild(m_AddressCount, L"AddressCount = 0x%1!08X!", m_AddressCount->getData());

			auto i = DWORD{};
			for (const auto& address : m_Addresses)
			{
				terminateBlock();
				addChild(
					address->PropertyCount,
					L"Addresses[%1!d!].PropertyCount = 0x%2!08X!\r\n",
					i,
					address->PropertyCount->getData());
				addChild(address->Pad, L"Addresses[%1!d!].Pad = 0x%2!08X!\r\n", i, address->Pad->getData());

				addHeader(L"Properties[%1!d!]:\r\n", i);
				addChild(address->Props.getBlock());
				i++;
			}
		}

		terminateBlock();
		addChild(m_SkipLen2, L"SkipLen2 = 0x%1!08X!", m_SkipLen2->getData());

		if (*m_SkipLen2)
		{
			terminateBlock();
			addHeader(L"SkipBytes2 = ");
			addChild(m_SkipBytes2);
		}

		if (m_Restriction && m_Restriction->hasData())
		{
			terminateBlock();
			addChild(m_Restriction->getBlock());
		}

		if (*m_Flags & SFST_FILTERSTREAM)
		{
			terminateBlock();
			addHeader(L"AdvancedSearchLen = 0x%1!08X!", m_AdvancedSearchBytes->size());

			if (!m_AdvancedSearchBytes->empty())
			{
				terminateBlock();
				addHeader(L"AdvancedSearchBytes = ");
				addChild(m_AdvancedSearchBytes);
			}
		}

		terminateBlock();
		addChild(m_SkipLen3, L"SkipLen3 = 0x%1!08X!", m_SkipLen3->getData());

		if (*m_SkipLen3)
		{
			terminateBlock();
			addHeader(L"SkipBytes3 = ");
			addChild(m_SkipBytes3);
		}
	}
} // namespace smartview