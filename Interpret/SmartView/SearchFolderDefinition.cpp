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

	_Check_return_ std::wstring SearchFolderDefinition::ToStringInternal()
	{
		auto szFlags = InterpretNumberAsStringProp(m_Flags, PR_WB_SF_STORAGE_TYPE);
		auto szSearchFolderDefinition = strings::formatmessage(
			IDS_SFDEFINITIONHEADER,
			m_Version.getData(),
			m_Flags.getData(),
			szFlags.c_str(),
			m_NumericSearch.getData(),
			m_TextSearchLength.getData());

		if (m_TextSearchLength)
		{
			szSearchFolderDefinition +=
				strings::formatmessage(IDS_SFDEFINITIONTEXTSEARCH, m_TextSearchLengthExtended.getData());

			if (!m_TextSearch.empty())
			{
				szSearchFolderDefinition += m_TextSearch;
			}
		}

		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPLEN1, m_SkipLen1.getData());

		if (m_SkipLen1)
		{
			szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPBYTES1);
			szSearchFolderDefinition += strings::BinToHexString(m_SkipBytes1.getData(), true);
		}

		szSearchFolderDefinition +=
			strings::formatmessage(IDS_SFDEFINITIONDEEPSEARCH, m_DeepSearch.getData(), m_FolderList1Length.getData());

		if (m_FolderList1Length)
		{
			szSearchFolderDefinition +=
				strings::formatmessage(IDS_SFDEFINITIONFOLDERLIST1, m_FolderList1LengthExtended.getData());

			if (!m_FolderList1.empty())
			{
				szSearchFolderDefinition += m_FolderList1;
			}
		}

		szSearchFolderDefinition +=
			strings::formatmessage(IDS_SFDEFINITIONFOLDERLISTLENGTH2, m_FolderList2Length.getData());

		if (m_FolderList2Length)
		{
			szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONFOLDERLIST2);
			szSearchFolderDefinition += m_FolderList2.ToString();
		}

		if (SFST_BINARY & m_Flags)
		{
			szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONADDRESSCOUNT, m_AddressCount.getData());

			for (DWORD i = 0; i < m_Addresses.size(); i++)
			{
				szSearchFolderDefinition += strings::formatmessage(
					IDS_SFDEFINITIONADDRESSES,
					i,
					m_Addresses[i].PropertyCount.getData(),
					i,
					m_Addresses[i].Pad.getData());

				szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONPROPERTIES, i);
				// TODO: Use the block instead
				szSearchFolderDefinition += m_Addresses[i].Props.ToString();
			}
		}

		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPLEN2, m_SkipLen2.getData());

		if (m_SkipLen2)
		{
			szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPBYTES2);
			szSearchFolderDefinition += strings::BinToHexString(m_SkipBytes2.getData(), true);
		}

		if (!m_Restriction.empty())
		{
			szSearchFolderDefinition += L"\r\n"; // STRING_OK
			// TODO: Use a proper block here
			szSearchFolderDefinition += m_Restriction;
		}

		if (SFST_FILTERSTREAM & m_Flags)
		{
			szSearchFolderDefinition +=
				strings::formatmessage(IDS_SFDEFINITIONADVANCEDSEARCHLEN, m_AdvancedSearchBytes.size());

			if (!m_AdvancedSearchBytes.empty())
			{
				szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONADVANCEDSEARCHBYTES);
				szSearchFolderDefinition += strings::BinToHexString(m_AdvancedSearchBytes, true);
			}
		}

		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPLEN3, m_SkipLen3.getData());

		if (m_SkipLen3)
		{
			szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPBYTES3);
			szSearchFolderDefinition += strings::BinToHexString(m_SkipBytes3, true);
		}

		return szSearchFolderDefinition;
	}
}