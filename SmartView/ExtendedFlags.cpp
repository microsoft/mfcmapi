#include "stdafx.h"
#include "..\stdafx.h"
#include "ExtendedFlags.h"
#include "..\String.h"
#include "..\ParseProperty.h"
#include "..\InterpretProp2.h"
#include "..\ExtraPropTags.h"

ExtendedFlags::ExtendedFlags(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_ulNumFlags = 0;
	m_pefExtendedFlags = NULL;
}

ExtendedFlags::~ExtendedFlags()
{
	ULONG i = 0;
	if (m_ulNumFlags && m_pefExtendedFlags)
	{
		for (i = 0; i < m_ulNumFlags; i++)
		{
			delete[] m_pefExtendedFlags[i].lpUnknownData;
		}
	}

	delete[] m_pefExtendedFlags;
}

void ExtendedFlags::Parse()
{
	// Run through the parser once to count the number of flag structs
	for (;;)
	{
		// Must have at least 2 bytes left to have another flag
		if (m_Parser.RemainingBytes() < 2) break;
		BYTE ulId = NULL;
		BYTE cbData = NULL;
		m_Parser.GetBYTE(&ulId);
		m_Parser.GetBYTE(&cbData);
		// Must have at least cbData bytes left to be a valid flag
		if (m_Parser.RemainingBytes() < cbData) break;

		m_Parser.Advance(cbData);
		m_ulNumFlags++;
	}

	// Now we parse for real
	m_Parser.Rewind();

	if (m_ulNumFlags && m_ulNumFlags < _MaxEntriesSmall)
		m_pefExtendedFlags = new ExtendedFlag[m_ulNumFlags];
	if (m_pefExtendedFlags)
	{
		memset(m_pefExtendedFlags, 0, sizeof(ExtendedFlag)*m_ulNumFlags);
		ULONG i = 0;
		bool bBadData = false;

		for (i = 0; i < m_ulNumFlags; i++)
		{
			m_Parser.GetBYTE(&m_pefExtendedFlags[i].Id);
			m_Parser.GetBYTE(&m_pefExtendedFlags[i].Cb);

			// If the structure says there's more bytes than remaining buffer, we're done parsing.
			if (m_Parser.RemainingBytes() < m_pefExtendedFlags[i].Cb)
			{
				m_ulNumFlags = i;
				break;
			}

			switch (m_pefExtendedFlags[i].Id)
			{
			case EFPB_FLAGS:
				if (m_pefExtendedFlags[i].Cb == sizeof(DWORD))
					m_Parser.GetDWORD(&m_pefExtendedFlags[i].Data.ExtendedFlags);
				else
					bBadData = true;
				break;
			case EFPB_CLSIDID:
				if (m_pefExtendedFlags[i].Cb == sizeof(GUID))
					m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), (LPBYTE)&m_pefExtendedFlags[i].Data.SearchFolderID);
				else
					bBadData = true;
				break;
			case EFPB_SFTAG:
				if (m_pefExtendedFlags[i].Cb == sizeof(DWORD))
					m_Parser.GetDWORD(&m_pefExtendedFlags[i].Data.SearchFolderTag);
				else
					bBadData = true;
				break;
			case EFPB_TODO_VERSION:
				if (m_pefExtendedFlags[i].Cb == sizeof(DWORD))
					m_Parser.GetDWORD(&m_pefExtendedFlags[i].Data.ToDoFolderVersion);
				else
					bBadData = true;
				break;
			default:
				m_Parser.GetBYTES(m_pefExtendedFlags[i].Cb, _MaxBytes, &m_pefExtendedFlags[i].lpUnknownData);
				break;
			}

			// If we encountered a bad flag, stop parsing
			if (bBadData)
			{
				m_ulNumFlags = i;
				break;
			}
		}
	}
}

_Check_return_ wstring ExtendedFlags::ToStringInternal()
{
	wstring szExtendedFlags = formatmessage(IDS_EXTENDEDFLAGSHEADER, m_ulNumFlags);

	if (m_ulNumFlags && m_pefExtendedFlags)
	{
		ULONG i = 0;
		for (i = 0; i < m_ulNumFlags; i++)
		{
			wstring szFlags = InterpretFlags(flagExtendedFolderFlagType, m_pefExtendedFlags[i].Id);
			szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGID,
				m_pefExtendedFlags[i].Id, szFlags.c_str(),
				m_pefExtendedFlags[i].Cb);

			switch (m_pefExtendedFlags[i].Id)
			{
			case EFPB_FLAGS:
				szFlags = InterpretFlags(flagExtendedFolderFlag, m_pefExtendedFlags[i].Data.ExtendedFlags);
				szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGDATAFLAG, m_pefExtendedFlags[i].Data.ExtendedFlags, szFlags.c_str());
				break;
			case EFPB_CLSIDID:
				szFlags = GUIDToString(&m_pefExtendedFlags[i].Data.SearchFolderID);
				szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGDATASFID, szFlags.c_str());
				break;
			case EFPB_SFTAG:
				szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGDATASFTAG,
					m_pefExtendedFlags[i].Data.SearchFolderTag);
				break;
			case EFPB_TODO_VERSION:
				szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGDATATODOVERSION, m_pefExtendedFlags[i].Data.ToDoFolderVersion);
				break;
			}

			if (m_pefExtendedFlags[i].lpUnknownData)
			{
				SBinary sBin = { 0 };
				szExtendedFlags += loadstring(IDS_EXTENDEDFLAGUNKNOWN);
				sBin.cb = m_pefExtendedFlags[i].Cb;
				sBin.lpb = m_pefExtendedFlags[i].lpUnknownData;
				szExtendedFlags += BinToHexString(&sBin, true);
			}
		}
	}

	return szExtendedFlags;
}