#include "stdafx.h"
#include "SearchFolderDefinition.h"
#include "RestrictionStruct.h"
#include "PropertyStruct.h"
#include <Interpret/String.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/SmartView/SmartView.h>

SearchFolderDefinition::SearchFolderDefinition()
{
	m_Version = 0;
	m_Flags = 0;
	m_NumericSearch = 0;
	m_TextSearchLength = 0;
	m_TextSearchLengthExtended = 0;
	m_SkipLen1 = 0;
	m_DeepSearch = 0;
	m_FolderList1Length = 0;
	m_FolderList1LengthExtended = 0;
	m_FolderList2Length = 0;
	m_AddressCount = 0;
	m_SkipLen2 = 0;
	m_SkipLen3 = 0;
}

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
		m_FolderList2.Init(cbRemainingBytes, m_Parser.GetCurrentAddress());

		m_Parser.Advance(cbRemainingBytes);
	}

	if (SFST_BINARY & m_Flags)
	{
		m_AddressCount = m_Parser.Get<DWORD>();
		if (m_AddressCount && m_AddressCount < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_AddressCount; i++)
			{
				AddressListEntryStruct addressListEntryStruct;
				addressListEntryStruct.PropertyCount = m_Parser.Get<DWORD>();
				addressListEntryStruct.Pad = m_Parser.Get<DWORD>();
				if (addressListEntryStruct.PropertyCount)
				{
					addressListEntryStruct.Props = BinToSPropValue(
						addressListEntryStruct.PropertyCount,
						false);
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
		res.Init(false, true);
		res.SmartViewParser::Init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
		res.DisableJunkParsing();
		m_Restriction = res.ToString();
		m_Parser.Advance(res.GetCurrentOffset());
	}

	if (SFST_FILTERSTREAM & m_Flags)
	{
		auto cbRemainingBytes = m_Parser.RemainingBytes();
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

_Check_return_ std::wstring SearchFolderDefinition::ToStringInternal()
{
	std::wstring szSearchFolderDefinition;

	auto szFlags = InterpretNumberAsStringProp(m_Flags, PR_WB_SF_STORAGE_TYPE);
	szSearchFolderDefinition = strings::formatmessage(IDS_SFDEFINITIONHEADER,
		m_Version,
		m_Flags,
		szFlags.c_str(),
		m_NumericSearch,
		m_TextSearchLength);

	if (m_TextSearchLength)
	{
		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONTEXTSEARCH,
			m_TextSearchLengthExtended);

		if (m_TextSearch.size())
		{
			szSearchFolderDefinition += m_TextSearch;
		}
	}

	szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPLEN1,
		m_SkipLen1);

	if (m_SkipLen1)
	{
		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPBYTES1);
		szSearchFolderDefinition += strings::BinToHexString(m_SkipBytes1, true);
	}

	szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONDEEPSEARCH,
		m_DeepSearch,
		m_FolderList1Length);

	if (m_FolderList1Length)
	{
		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONFOLDERLIST1,
			m_FolderList1LengthExtended);

		if (m_FolderList1.size())
		{
			szSearchFolderDefinition += m_FolderList1;
		}
	}

	szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONFOLDERLISTLENGTH2,
		m_FolderList2Length);

	if (m_FolderList2Length)
	{
		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONFOLDERLIST2);
		szSearchFolderDefinition += m_FolderList2.ToString();
	}

	if (SFST_BINARY & m_Flags)
	{
		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONADDRESSCOUNT,
			m_AddressCount);

		for (DWORD i = 0; i < m_Addresses.size(); i++)
		{
			szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONADDRESSES,
				i, m_Addresses[i].PropertyCount,
				i, m_Addresses[i].Pad);

			szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONPROPERTIES, i);
			szSearchFolderDefinition += PropsToString(m_Addresses[i].PropertyCount, m_Addresses[i].Props);
		}
	}

	szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPLEN2,
		m_SkipLen2);

	if (m_SkipLen2)
	{
		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPBYTES2);
		szSearchFolderDefinition += strings::BinToHexString(m_SkipBytes2, true);
	}

	if (!m_Restriction.empty())
	{
		szSearchFolderDefinition += L"\r\n"; // STRING_OK
		szSearchFolderDefinition += m_Restriction;
	}

	if (SFST_FILTERSTREAM & m_Flags)
	{
		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONADVANCEDSEARCHLEN,
			m_AdvancedSearchBytes.size());

		if (m_AdvancedSearchBytes.size())
		{
			szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONADVANCEDSEARCHBYTES);
			szSearchFolderDefinition += strings::BinToHexString(m_AdvancedSearchBytes, true);
		}
	}

	szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPLEN3,
		m_SkipLen3);

	if (m_SkipLen3)
	{
		szSearchFolderDefinition += strings::formatmessage(IDS_SFDEFINITIONSKIPBYTES3);
		szSearchFolderDefinition += strings::BinToHexString(m_SkipBytes3, true);
	}

	return szSearchFolderDefinition;
}