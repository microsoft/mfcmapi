#include <core/stdafx.h>
#include <core/smartview/SearchFolderDefinition.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>

namespace smartview
{
	void AddressListEntryStruct::parse()
	{
		PropertyCount = blockT<DWORD>::parse(parser);
		Pad = blockT<DWORD>::parse(parser);
		if (*PropertyCount)
		{
			Props = std::make_shared<PropertiesStruct>(*PropertyCount, false, false);
			Props->block::parse(parser, false);
		}
	}

	void AddressListEntryStruct::parseBlocks()
	{
		addChild(PropertyCount, L"PropertyCount = 0x%1!08X!", PropertyCount->getData());
		addChild(Pad, L"Pad = 0x%1!08X!", Pad->getData());
		addChild(Props, L"Properties");
	}

	void SearchFolderDefinition::parse()
	{
		m_Version = blockT<DWORD>::parse(parser);
		m_Flags = blockT<DWORD>::parse(parser);
		m_NumericSearch = blockT<DWORD>::parse(parser);

		m_TextSearchLength = blockT<BYTE>::parse(parser);
		size_t cchTextSearch = *m_TextSearchLength;
		if (*m_TextSearchLength == 255)
		{
			m_TextSearchLengthExtended = blockT<WORD>::parse(parser);
			cchTextSearch = *m_TextSearchLengthExtended;
		}

		if (cchTextSearch)
		{
			m_TextSearch = blockStringW::parse(parser, cchTextSearch);
		}

		m_SkipLen1 = blockT<DWORD>::parse(parser);
		m_SkipBytes1 = blockBytes::parse(parser, *m_SkipLen1, _MaxBytes);

		m_DeepSearch = blockT<DWORD>::parse(parser);

		m_FolderList1Length = blockT<BYTE>::parse(parser);
		size_t cchFolderList1 = *m_FolderList1Length;
		if (*m_FolderList1Length == 255)
		{
			m_FolderList1LengthExtended = blockT<WORD>::parse(parser);
			cchFolderList1 = *m_FolderList1LengthExtended;
		}

		if (cchFolderList1)
		{
			m_FolderList1 = blockStringW::parse(parser, cchFolderList1);
		}

		m_FolderList2Length = blockT<DWORD>::parse(parser);

		if (m_FolderList2Length)
		{
			m_FolderList2 = block::parse<EntryList>(parser, *m_FolderList2Length, true);
		}

		if (*m_Flags & SFST_BINARY)
		{
			m_AddressCount = blockT<DWORD>::parse(parser);
			if (*m_AddressCount)
			{
				if (*m_AddressCount < _MaxEntriesSmall)
				{
					m_Addresses.reserve(*m_AddressCount);
					for (DWORD i = 0; i < *m_AddressCount; i++)
					{
						m_Addresses.emplace_back(block::parse<AddressListEntryStruct>(parser, false));
					}
				}
			}
		}

		m_SkipLen2 = blockT<DWORD>::parse(parser);
		m_SkipBytes2 = blockBytes::parse(parser, *m_SkipLen2, _MaxBytes);

		if (*m_Flags & SFST_MRES)
		{
			m_Restriction = std::make_shared<RestrictionStruct>(false, true);
			m_Restriction->block::parse(parser, false);
		}

		if (*m_Flags & SFST_FILTERSTREAM)
		{
			const auto cbRemainingBytes = parser->getSize();
			// Since the format for SFST_FILTERSTREAM isn't documented, just assume that everything remaining
			// is part of this bucket. We leave DWORD space for the final skip block, which should be empty
			if (cbRemainingBytes > sizeof DWORD)
			{
				m_AdvancedSearchBytes = blockBytes::parse(parser, cbRemainingBytes - sizeof DWORD);
			}
		}

		m_SkipLen3 = blockT<DWORD>::parse(parser);
		if (m_SkipLen3)
		{
			m_SkipBytes3 = blockBytes::parse(parser, *m_SkipLen3, _MaxBytes);
		}
	}

	void SearchFolderDefinition::parseBlocks()
	{
		setText(L"Search Folder Definition");
		addChild(m_Version, L"Version = 0x%1!08X!", m_Version->getData());
		addChild(
			m_Flags,
			L"Flags = 0x%1!08X! = %2!ws!",
			m_Flags->getData(),
			InterpretNumberAsStringProp(*m_Flags, PR_WB_SF_STORAGE_TYPE).c_str());
		addChild(m_NumericSearch, L"Numeric Search = 0x%1!08X!", m_NumericSearch->getData());
		addChild(m_TextSearchLength, L"Text Search Length = 0x%1!02X!", m_TextSearchLength->getData());

		if (*m_TextSearchLength)
		{
			addChild(
				m_TextSearchLengthExtended,
				L"Text Search Length Extended = 0x%1!04X!",
				m_TextSearchLengthExtended->getData());
			addLabeledChild(L"Text Search", m_TextSearch);
		}

		addChild(m_SkipLen1, L"SkipLen1 = 0x%1!08X!", m_SkipLen1->getData());

		if (*m_SkipLen1)
		{
			addLabeledChild(L"SkipBytes1", m_SkipBytes1);
		}

		addChild(m_DeepSearch, L"Deep Search = 0x%1!08X!", m_DeepSearch->getData());
		addChild(m_FolderList1Length, L"Folder List 1 Length = 0x%1!02X!", m_FolderList1Length->getData());

		if (*m_FolderList1Length)
		{
			addChild(
				m_FolderList1LengthExtended,
				L"Folder List 1 Length Extended = 0x%1!04X!",
				m_FolderList1LengthExtended->getData());
			addLabeledChild(L"Folder List 1", m_FolderList1);
		}

		addChild(m_FolderList2Length, L"Folder List 2 Length = 0x%1!08X!", m_FolderList2Length->getData());

		if (m_FolderList2)
		{
			addLabeledChild(L"Folder List2", m_FolderList2);
		}

		if (*m_Flags & SFST_BINARY)
		{
			addChild(m_AddressCount, L"AddressCount = 0x%1!08X!", m_AddressCount->getData());

			auto i = DWORD{};
			for (const auto& address : m_Addresses)
			{
				addChild(address, L"Addresses[%1!d!]", i);
				i++;
			}
		}

		addChild(m_SkipLen2, L"SkipLen2 = 0x%1!08X!", m_SkipLen2->getData());

		addLabeledChild(L"SkipBytes2", m_SkipBytes2);

		if (m_Restriction && m_Restriction->hasData())
		{
			addChild(m_Restriction);
		}

		if (*m_Flags & SFST_FILTERSTREAM)
		{
			addHeader(L"AdvancedSearchLen = 0x%1!08X!", m_AdvancedSearchBytes->size());

			if (!m_AdvancedSearchBytes->empty())
			{
				addLabeledChild(L"AdvancedSearchBytes", m_AdvancedSearchBytes);
			}
		}

		addChild(m_SkipLen3, L"SkipLen3 = 0x%1!08X!", m_SkipLen3->getData());

		addLabeledChild(L"SkipBytes3", m_SkipBytes3);
	}
} // namespace smartview