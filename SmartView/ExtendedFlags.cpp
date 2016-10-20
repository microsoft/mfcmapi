#include "stdafx.h"
#include "ExtendedFlags.h"
#include "String.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

ExtendedFlags::ExtendedFlags()
{
	m_ulNumFlags = 0;
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
	{
		auto bBadData = false;

		for (ULONG i = 0; i < m_ulNumFlags; i++)
		{
			ExtendedFlag extendedFlag;
			m_pefExtendedFlags.push_back(extendedFlag);

			m_Parser.GetBYTE(&extendedFlag.Id);
			m_Parser.GetBYTE(&extendedFlag.Cb);

			// If the structure says there's more bytes than remaining buffer, we're done parsing.
			if (m_Parser.RemainingBytes() < extendedFlag.Cb)
			{
				m_ulNumFlags = i;
				break;
			}

			switch (extendedFlag.Id)
			{
			case EFPB_FLAGS:
				if (extendedFlag.Cb == sizeof(DWORD))
					m_Parser.GetDWORD(&extendedFlag.Data.ExtendedFlags);
				else
					bBadData = true;
				break;
			case EFPB_CLSIDID:
				if (extendedFlag.Cb == sizeof(GUID))
					m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), reinterpret_cast<LPBYTE>(&extendedFlag.Data.SearchFolderID));
				else
					bBadData = true;
				break;
			case EFPB_SFTAG:
				if (extendedFlag.Cb == sizeof(DWORD))
					m_Parser.GetDWORD(&extendedFlag.Data.SearchFolderTag);
				else
					bBadData = true;
				break;
			case EFPB_TODO_VERSION:
				if (extendedFlag.Cb == sizeof(DWORD))
					m_Parser.GetDWORD(&extendedFlag.Data.ToDoFolderVersion);
				else
					bBadData = true;
				break;
			default:
				extendedFlag.lpUnknownData = m_Parser.GetBYTES(extendedFlag.Cb, _MaxBytes);
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
	auto szExtendedFlags = formatmessage(IDS_EXTENDEDFLAGSHEADER, m_ulNumFlags);

	if (m_pefExtendedFlags.size())
	{
		for (auto extendedFlag : m_pefExtendedFlags)
		{
			auto szFlags = InterpretFlags(flagExtendedFolderFlagType, extendedFlag.Id);
			szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGID,
				extendedFlag.Id, szFlags.c_str(),
				extendedFlag.Cb);

			switch (extendedFlag.Id)
			{
			case EFPB_FLAGS:
				szFlags = InterpretFlags(flagExtendedFolderFlag, extendedFlag.Data.ExtendedFlags);
				szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGDATAFLAG, extendedFlag.Data.ExtendedFlags, szFlags.c_str());
				break;
			case EFPB_CLSIDID:
				szFlags = GUIDToString(&extendedFlag.Data.SearchFolderID);
				szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGDATASFID, szFlags.c_str());
				break;
			case EFPB_SFTAG:
				szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGDATASFTAG,
					extendedFlag.Data.SearchFolderTag);
				break;
			case EFPB_TODO_VERSION:
				szExtendedFlags += formatmessage(IDS_EXTENDEDFLAGDATATODOVERSION, extendedFlag.Data.ToDoFolderVersion);
				break;
			}

			if (extendedFlag.lpUnknownData.size())
			{
				szExtendedFlags += loadstring(IDS_EXTENDEDFLAGUNKNOWN);
				szExtendedFlags += BinToHexString(extendedFlag.lpUnknownData, true);
			}
		}
	}

	return szExtendedFlags;
}