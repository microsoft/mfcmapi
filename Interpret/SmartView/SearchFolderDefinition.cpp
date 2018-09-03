#include <StdAfx.h>
#include <Interpret/SmartView/SearchFolderDefinition.h>
#include <Interpret/SmartView/RestrictionStruct.h>
#include <Interpret/SmartView/PropertyStruct.h>
#include <Interpret/String.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/SmartView/SmartView.h>

namespace smartview
{
	SearchFolderDefinition::SearchFolderDefinition() {}

	void SearchFolderDefinition::Parse()
	{
		m_Version = m_Parser.GetBlock<DWORD>();
		m_Flags = m_Parser.GetBlock<DWORD>();
		m_NumericSearch = m_Parser.GetBlock<DWORD>();

		m_TextSearchLength = m_Parser.GetBlock<BYTE>();
		size_t cchTextSearch = m_TextSearchLength;
		if (255 == m_TextSearchLength)
		{
			m_TextSearchLengthExtended = m_Parser.GetBlock<WORD>();
			cchTextSearch = m_TextSearchLengthExtended;
		}

		if (cchTextSearch)
		{
			m_TextSearch = m_Parser.GetBlockStringW(cchTextSearch);
		}

		m_SkipLen1 = m_Parser.GetBlock<DWORD>();
		m_SkipBytes1 = m_Parser.GetBlockBYTES(m_SkipLen1, _MaxBytes);

		m_DeepSearch = m_Parser.GetBlock<DWORD>();

		m_FolderList1Length = m_Parser.GetBlock<BYTE>();
		size_t cchFolderList1 = m_FolderList1Length;
		if (255 == m_FolderList1Length)
		{
			m_FolderList1LengthExtended = m_Parser.GetBlock<WORD>();
			cchFolderList1 = m_FolderList1LengthExtended;
		}

		if (cchFolderList1)
		{
			m_FolderList1 = m_Parser.GetBlockStringW(cchFolderList1);
		}

		m_FolderList2Length = m_Parser.GetBlock<DWORD>();

		if (m_FolderList2Length)
		{
			auto cbRemainingBytes = m_Parser.RemainingBytes();
			cbRemainingBytes = min(m_FolderList2Length, cbRemainingBytes);
			m_FolderList2.Init(cbRemainingBytes, m_Parser.GetCurrentAddress());
			m_FolderList2.EnsureParsed();

			m_Parser.Advance(cbRemainingBytes);
		}

		if (SFST_BINARY & m_Flags)
		{
			m_AddressCount = m_Parser.GetBlock<DWORD>();
			if (m_AddressCount && m_AddressCount < _MaxEntriesSmall)
			{
				for (DWORD i = 0; i < m_AddressCount; i++)
				{
					AddressListEntryStruct addressListEntryStruct{};
					addressListEntryStruct.PropertyCount = m_Parser.GetBlock<DWORD>();
					addressListEntryStruct.Pad = m_Parser.GetBlock<DWORD>();
					if (addressListEntryStruct.PropertyCount)
					{
						addressListEntryStruct.Props.Init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
						addressListEntryStruct.Props.SetMaxEntries(addressListEntryStruct.PropertyCount);
						addressListEntryStruct.Props.DisableJunkParsing();
						addressListEntryStruct.Props.EnsureParsed();
						m_Parser.Advance(addressListEntryStruct.Props.GetCurrentOffset());
					}

					m_Addresses.push_back(addressListEntryStruct);
				}
			}
		}

		m_SkipLen2 = m_Parser.GetBlock<DWORD>();
		m_SkipBytes2 = m_Parser.GetBlockBYTES(m_SkipLen2, _MaxBytes);

		if (SFST_MRES & m_Flags)
		{
			RestrictionStruct res;
			res.Init(false, true);
			res.SmartViewParser::Init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
			res.DisableJunkParsing();
			m_Restriction = blockStringW{};
			// TODO: Make this a proper block
			m_Restriction.setData(res.ToString());
			m_Parser.Advance(res.GetCurrentOffset());
		}

		if (SFST_FILTERSTREAM & m_Flags)
		{
			const auto cbRemainingBytes = m_Parser.RemainingBytes();
			// Since the format for SFST_FILTERSTREAM isn't documented, just assume that everything remaining
			// is part of this bucket. We leave DWORD space for the final skip block, which should be empty
			if (cbRemainingBytes > sizeof DWORD)
			{
				m_AdvancedSearchBytes = m_Parser.GetBlockBYTES(cbRemainingBytes - sizeof DWORD);
			}
		}

		m_SkipLen3 = m_Parser.GetBlock<DWORD>();
		if (m_SkipLen3)
		{
			m_SkipBytes3 = m_Parser.GetBlockBYTES(m_SkipLen3, _MaxBytes);
		}
	}

	void SearchFolderDefinition::ParseBlocks()
	{
		addHeader(L"Search Folder Definition:\r\n");
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
			addLine();
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

		addLine();
		addBlock(m_SkipLen1, L"SkipLen1 = 0x%1!08X!", m_SkipLen1.getData());

		if (m_SkipLen1)
		{
			addLine();
			addHeader(L"SkipBytes1 = ");

			addBlockBytes(m_SkipBytes1);
		}

		addLine();
		addBlock(m_DeepSearch, L"Deep Search = 0x%1!08X!\r\n", m_DeepSearch.getData());
		addBlock(m_FolderList1Length, L"Folder List 1 Length = 0x%1!02X!", m_FolderList1Length.getData());

		if (m_FolderList1Length)
		{
			addLine();
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

		addLine();
		addBlock(m_FolderList2Length, L"Folder List 2 Length = 0x%1!08X!", m_FolderList2Length.getData());

		if (m_FolderList2Length)
		{
			addLine();
			addHeader(L"FolderList2 = \r\n");
			addBlock(m_FolderList2.getBlock());
		}

		if (SFST_BINARY & m_Flags)
		{
			addLine();
			addBlock(m_AddressCount, L"AddressCount = 0x%1!08X!", m_AddressCount.getData());

			for (DWORD i = 0; i < m_Addresses.size(); i++)
			{
				addLine();
				addBlock(
					m_Addresses[i].PropertyCount,
					L"Addresses[%1!d!].PropertyCount = 0x%2!08X!\r\n",
					i,
					m_Addresses[i].PropertyCount.getData());
				addBlock(m_Addresses[i].Pad, L"Addresses[%1!d!].Pad = 0x%2!08X!", i, m_Addresses[i].Pad.getData());

				addLine();
				addHeader(L"Properties[%1!d!]:\r\n", i);
				addBlock(m_Addresses[i].Props.getBlock());
			}
		}

		addLine();
		addBlock(m_SkipLen2, L"SkipLen2 = 0x%1!08X!", m_SkipLen2.getData());

		if (m_SkipLen2)
		{
			addLine();
			addHeader(L"SkipBytes2 = ");
			addBlockBytes(m_SkipBytes2);
		}

		if (!m_Restriction.empty())
		{
			addLine();
			// TODO: Use a proper block here
			addBlock(m_Restriction, m_Restriction.c_str());
		}

		if (SFST_FILTERSTREAM & m_Flags)
		{
			addLine();
			addBlock(m_AdvancedSearchBytes, L"AdvancedSearchLen = 0x%1!08X!", m_AdvancedSearchBytes.size());

			if (!m_AdvancedSearchBytes.empty())
			{
				addLine();
				addHeader(L"AdvancedSearchBytes = ");
				addBlockBytes(m_AdvancedSearchBytes);
				addLine();
			}
		}

		addBlock(m_SkipLen3, L"SkipLen3 = 0x%1!08X!", m_SkipLen3.getData());

		if (m_SkipLen3)
		{
			addLine();
			addHeader(L"SkipBytes3 = ");
			addBlockBytes(m_SkipBytes3);
		}
	}
}