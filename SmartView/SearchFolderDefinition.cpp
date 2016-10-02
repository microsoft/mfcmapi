#include "stdafx.h"
#include "SearchFolderDefinition.h"
#include "RestrictionStruct.h"
#include "PropertyStruct.h"
#include "String.h"
#include "ExtraPropTags.h"

SearchFolderDefinition::SearchFolderDefinition(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_Version = 0;
	m_Flags = 0;
	m_NumericSearch = 0;
	m_TextSearchLength = 0;
	m_TextSearchLengthExtended = 0;
	m_TextSearch = nullptr;
	m_SkipLen1 = 0;
	m_SkipBytes1 = nullptr;
	m_DeepSearch = 0;
	m_FolderList1Length = 0;
	m_FolderList1LengthExtended = 0;
	m_FolderList1 = nullptr;
	m_FolderList2Length = 0;
	m_FolderList2 = nullptr;
	m_AddressCount = 0;
	m_Addresses = nullptr;
	m_SkipLen2 = 0;
	m_SkipBytes2 = nullptr;
	m_AdvancedSearchLen = 0;
	m_AdvancedSearchBytes = nullptr;
	m_SkipLen3 = 0;
	m_SkipBytes3 = nullptr;
}

SearchFolderDefinition::~SearchFolderDefinition()
{
	delete[] m_TextSearch;
	delete[] m_SkipBytes1;
	delete[] m_FolderList1;
	if (m_FolderList2) delete m_FolderList2;
	if (m_Addresses)
	{
		for (DWORD i = 0; i < m_AddressCount; i++)
		{
			DeleteSPropVal(m_Addresses[i].PropertyCount, m_Addresses[i].Props);
		}

		delete[] m_Addresses;
	}

	delete[] m_SkipBytes2;
	delete[] m_AdvancedSearchBytes;
	delete[] m_SkipBytes3;
}

void SearchFolderDefinition::Parse()
{
	m_Parser.GetDWORD(&m_Version);
	m_Parser.GetDWORD(&m_Flags);
	m_Parser.GetDWORD(&m_NumericSearch);

	m_Parser.GetBYTE(&m_TextSearchLength);
	size_t cchTextSearch = m_TextSearchLength;
	if (255 == m_TextSearchLength)
	{
		m_Parser.GetWORD(&m_TextSearchLengthExtended);
		cchTextSearch = m_TextSearchLengthExtended;
	}

	if (cchTextSearch)
	{
		m_Parser.GetStringW(cchTextSearch, &m_TextSearch);
	}

	m_Parser.GetDWORD(&m_SkipLen1);
	m_Parser.GetBYTES(m_SkipLen1, _MaxBytes, &m_SkipBytes1);

	m_Parser.GetDWORD(&m_DeepSearch);

	m_Parser.GetBYTE(&m_FolderList1Length);
	size_t cchFolderList1 = m_FolderList1Length;
	if (255 == m_FolderList1Length)
	{
		m_Parser.GetWORD(&m_FolderList1LengthExtended);
		cchFolderList1 = m_FolderList1LengthExtended;
	}

	if (cchFolderList1)
	{
		m_Parser.GetStringW(cchFolderList1, &m_FolderList1);
	}

	m_Parser.GetDWORD(&m_FolderList2Length);

	if (m_FolderList2Length)
	{
		auto cbRemainingBytes = m_Parser.RemainingBytes();
		cbRemainingBytes = min(m_FolderList2Length, cbRemainingBytes);
		m_FolderList2 = new EntryList(
			static_cast<ULONG>(cbRemainingBytes),
			m_Parser.GetCurrentAddress());
		m_Parser.Advance(cbRemainingBytes);
	}

	if (SFST_BINARY & m_Flags)
	{
		m_Parser.GetDWORD(&m_AddressCount);
		if (m_AddressCount && m_AddressCount < _MaxEntriesSmall)
		{
			m_Addresses = new AddressListEntryStruct[m_AddressCount];

			if (m_Addresses)
			{
				memset(m_Addresses, 0, m_AddressCount * sizeof(AddressListEntryStruct));

				for (DWORD i = 0; i < m_AddressCount; i++)
				{
					m_Parser.GetDWORD(&m_Addresses[i].PropertyCount);
					m_Parser.GetDWORD(&m_Addresses[i].Pad);
					if (m_Addresses[i].PropertyCount)
					{
						m_Addresses[i].Props = BinToSPropValue(
							m_Addresses[i].PropertyCount,
							false);
					}
				}
			}
		}
	}

	m_Parser.GetDWORD(&m_SkipLen2);
	m_Parser.GetBYTES(m_SkipLen2, _MaxBytes, &m_SkipBytes2);

	if (SFST_MRES & m_Flags)
	{
		auto res = new RestrictionStruct(
			static_cast<ULONG>(m_Parser.RemainingBytes()),
			m_Parser.GetCurrentAddress(),
			false,
			true);
		if (res)
		{
			res->DisableJunkParsing();
			m_Restriction = res->ToString();
			m_Parser.Advance(res->GetCurrentOffset());
			delete res;
		}
	}

	if (SFST_FILTERSTREAM & m_Flags)
	{
		auto cbRemainingBytes = m_Parser.RemainingBytes();
		// Since the format for SFST_FILTERSTREAM isn't documented, just assume that everything remaining
		// is part of this bucket. We leave DWORD space for the final skip block, which should be empty
		if (cbRemainingBytes > sizeof(DWORD))
		{
			m_AdvancedSearchLen = static_cast<DWORD>(cbRemainingBytes) - sizeof(DWORD);
			m_Parser.GetBYTES(m_AdvancedSearchLen, m_AdvancedSearchLen, &m_AdvancedSearchBytes);
		}
	}

	m_Parser.GetDWORD(&m_SkipLen3);
	if (m_SkipLen3)
	{
		m_Parser.GetBYTES(m_SkipLen3, _MaxBytes, &m_SkipBytes3);
	}
}

_Check_return_ wstring SearchFolderDefinition::ToStringInternal()
{
	wstring szSearchFolderDefinition;

	auto szFlags = InterpretNumberAsStringProp(m_Flags, PR_WB_SF_STORAGE_TYPE);
	szSearchFolderDefinition = formatmessage(IDS_SFDEFINITIONHEADER,
		m_Version,
		m_Flags,
		szFlags.c_str(),
		m_NumericSearch,
		m_TextSearchLength);

	if (m_TextSearchLength)
	{
		szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONTEXTSEARCH,
			m_TextSearchLengthExtended);

		if (m_TextSearch)
		{
			szSearchFolderDefinition += m_TextSearch;
		}
	}

	szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONSKIPLEN1,
		m_SkipLen1);

	if (m_SkipLen1)
	{
		SBinary sBin = { 0 };

		sBin.cb = static_cast<ULONG>(m_SkipLen1);
		sBin.lpb = m_SkipBytes1;

		szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONSKIPBYTES1);
		szSearchFolderDefinition += BinToHexString(&sBin, true);
	}

	szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONDEEPSEARCH,
		m_DeepSearch,
		m_FolderList1Length);

	if (m_FolderList1Length)
	{
		szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONFOLDERLIST1,
			m_FolderList1LengthExtended);

		if (m_FolderList1)
		{
			szSearchFolderDefinition += m_FolderList1;
		}
	}

	szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONFOLDERLISTLENGTH2,
		m_FolderList2Length);

	if (m_FolderList2Length && m_FolderList2)
	{
		szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONFOLDERLIST2);
		szSearchFolderDefinition += m_FolderList2->ToString();
	}

	if (SFST_BINARY & m_Flags)
	{
		szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONADDRESSCOUNT,
			m_AddressCount);
		if (m_Addresses && m_AddressCount)
		{
			for (DWORD i = 0; i < m_AddressCount; i++)
			{
				szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONADDRESSES,
					i, m_Addresses[i].PropertyCount,
					i, m_Addresses[i].Pad);

				szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONPROPERTIES, i);
				szSearchFolderDefinition += PropsToString(m_Addresses[i].PropertyCount, m_Addresses[i].Props);
			}
		}
	}

	szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONSKIPLEN2,
		m_SkipLen2);

	if (m_SkipLen2)
	{
		SBinary sBin = { 0 };

		sBin.cb = static_cast<ULONG>(m_SkipLen2);
		sBin.lpb = m_SkipBytes2;

		szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONSKIPBYTES2);
		szSearchFolderDefinition += BinToHexString(&sBin, true);
	}

	if (!m_Restriction.empty())
	{
		szSearchFolderDefinition += L"\r\n"; // STRING_OK
		szSearchFolderDefinition += m_Restriction;
	}

	if (SFST_FILTERSTREAM & m_Flags)
	{
		szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONADVANCEDSEARCHLEN,
			m_AdvancedSearchLen);

		if (m_AdvancedSearchLen)
		{
			SBinary sBin = { 0 };

			sBin.cb = static_cast<ULONG>(m_AdvancedSearchLen);
			sBin.lpb = m_AdvancedSearchBytes;

			szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONADVANCEDSEARCHBYTES);
			szSearchFolderDefinition += BinToHexString(&sBin, true);
		}
	}

	szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONSKIPLEN3,
		m_SkipLen3);

	if (m_SkipLen3)
	{
		SBinary sBin = { 0 };

		sBin.cb = static_cast<ULONG>(m_SkipLen3);
		sBin.lpb = m_SkipBytes3;

		szSearchFolderDefinition += formatmessage(IDS_SFDEFINITIONSKIPBYTES3);
		szSearchFolderDefinition += BinToHexString(&sBin, true);
	}

	return szSearchFolderDefinition;
}