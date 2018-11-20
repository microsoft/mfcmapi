#include <StdAfx.h>
#include <Interpret/SmartView/SearchFolderDefinition.h>
#include <Interpret/SmartView/RestrictionStruct.h>
#include <Interpret/SmartView/PropertiesStruct.h>
#include <Interpret/String.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/SmartView/SmartView.h>

namespace smartview
{
	void SearchFolderDefinition::Parse()
	{
		m_Version = m_Parser.Get<DWORD>();
		m_Flags = m_Parser.Get<DWORD>();
		m_NumericSearch = m_Parser.Get<DWORD>();

		m_TextSearchLength = m_Parser.Get<BYTE>();
		size_t cchTextSearch = m_TextSearchLength;
		if (255 == m_TextSearchLength)
		{
			m_TextSearchLengthExtended = m_Parser.Get<WORD>();
			cchTextSearch = m_TextSearchLengthExtended;
		}

		if (cchTextSearch)
		{
			m_TextSearch = m_Parser.GetStringW(cchTextSearch);
		}

		m_SkipLen1 = m_Parser.Get<DWORD>();
		m_SkipBytes1 = m_Parser.GetBYTES(m_SkipLen1, _MaxBytes);

		m_DeepSearch = m_Parser.Get<DWORD>();

		m_FolderList1Length = m_Parser.Get<BYTE>();
		size_t cchFolderList1 = m_FolderList1Length;
		if (255 == m_FolderList1Length)
		{
			m_FolderList1LengthExtended = m_Parser.Get<WORD>();
			cchFolderList1 = m_FolderList1LengthExtended;
		}

		if (cchFolderList1)
		{
			m_FolderList1 = m_Parser.GetStringW(cchFolderList1);
		}

		m_FolderList2Length = m_Parser.Get<DWORD>();

		if (m_FolderList2Length)
		{
			auto cbRemainingBytes = m_Parser.RemainingBytes();
			cbRemainingBytes = min(m_FolderList2Length, cbRemainingBytes);
			m_FolderList2.parse(m_Parser, cbRemainingBytes, true);
		}

		if (SFST_BINARY & m_Flags)
		{
			m_AddressCount = m_Parser.Get<DWORD>();
			if (m_AddressCount && m_AddressCount < _MaxEntriesSmall)
			{
				m_Addresses.reserve(m_AddressCount);
				for (DWORD i = 0; i < m_AddressCount; i++)
				{
					AddressListEntryStruct addressListEntryStruct{};
					addressListEntryStruct.PropertyCount = m_Parser.Get<DWORD>();
					addressListEntryStruct.Pad = m_Parser.Get<DWORD>();
					if (addressListEntryStruct.PropertyCount)
					{
						addressListEntryStruct.Props.SetMaxEntries(addressListEntryStruct.PropertyCount);
						addressListEntryStruct.Props.parse(m_Parser, false);
					}

					m_Addresses.push_back(addressListEntryStruct);
				}
			}
		}

		m_SkipLen2 = m_Parser.Get<DWORD>();
		m_SkipBytes2 = m_Parser.GetBYTES(m_SkipLen2, _MaxBytes);

		if (SFST_MRES & m_Flags)
		{
			RestrictionStruct res;
			res.init(false, true);
			res.parse(m_Parser, false);
			m_Restriction = res.getBlock();
		}

		if (SFST_FILTERSTREAM & m_Flags)
		{
			const auto cbRemainingBytes = m_Parser.RemainingBytes();
			// Since the format for SFST_FILTERSTREAM isn't documented, just assume that everything remaining
			// is part of this bucket. We leave DWORD space for the final skip block, which should be empty
			if (cbRemainingBytes > sizeof DWORD)
			{
				m_AdvancedSearchBytes = m_Parser.GetBYTES(cbRemainingBytes - sizeof DWORD);
			}
		}

		m_SkipLen3 = m_Parser.Get<DWORD>();
		if (m_SkipLen3)
		{
			m_SkipBytes3 = m_Parser.GetBYTES(m_SkipLen3, _MaxBytes);
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

			for (DWORD i = 0; i < m_Addresses.size(); i++)
			{
				terminateBlock();
				addBlock(
					m_Addresses[i].PropertyCount,
					L"Addresses[%1!d!].PropertyCount = 0x%2!08X!\r\n",
					i,
					m_Addresses[i].PropertyCount.getData());
				addBlock(m_Addresses[i].Pad, L"Addresses[%1!d!].Pad = 0x%2!08X!\r\n", i, m_Addresses[i].Pad.getData());

				addHeader(L"Properties[%1!d!]:\r\n", i);
				addBlock(m_Addresses[i].Props.getBlock());
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

		if (m_Restriction.hasData())
		{
			terminateBlock();
			addBlock(m_Restriction);
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