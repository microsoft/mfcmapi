#include "stdafx.h"
#include "ExtendedFlags.h"
#include <Interpret/InterpretProp2.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
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
			(void)m_Parser.Get<BYTE>();
			const auto cbData = m_Parser.Get<BYTE>();
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

				extendedFlag.Id = m_Parser.Get<BYTE>();
				extendedFlag.Cb = m_Parser.Get<BYTE>();

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
						extendedFlag.Data.ExtendedFlags = m_Parser.Get<DWORD>();
					else
						bBadData = true;
					break;
				case EFPB_CLSIDID:
					if (extendedFlag.Cb == sizeof(GUID))
						extendedFlag.Data.SearchFolderID = m_Parser.Get<GUID>();
					else
						bBadData = true;
					break;
				case EFPB_SFTAG:
					if (extendedFlag.Cb == sizeof(DWORD))
						extendedFlag.Data.SearchFolderTag = m_Parser.Get<DWORD>();
					else
						bBadData = true;
					break;
				case EFPB_TODO_VERSION:
					if (extendedFlag.Cb == sizeof(DWORD))
						extendedFlag.Data.ToDoFolderVersion = m_Parser.Get<DWORD>();
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

				m_pefExtendedFlags.push_back(extendedFlag);
			}
		}
	}

	_Check_return_ std::wstring ExtendedFlags::ToStringInternal()
	{
		auto szExtendedFlags = strings::formatmessage(IDS_EXTENDEDFLAGSHEADER, m_ulNumFlags);

		if (m_pefExtendedFlags.size())
		{
			for (const auto& extendedFlag : m_pefExtendedFlags)
			{
				auto szFlags = InterpretFlags(flagExtendedFolderFlagType, extendedFlag.Id);
				szExtendedFlags += strings::formatmessage(IDS_EXTENDEDFLAGID,
					extendedFlag.Id, szFlags.c_str(),
					extendedFlag.Cb);

				switch (extendedFlag.Id)
				{
				case EFPB_FLAGS:
					szFlags = InterpretFlags(flagExtendedFolderFlag, extendedFlag.Data.ExtendedFlags);
					szExtendedFlags += strings::formatmessage(IDS_EXTENDEDFLAGDATAFLAG, extendedFlag.Data.ExtendedFlags, szFlags.c_str());
					break;
				case EFPB_CLSIDID:
					szFlags = GUIDToString(&extendedFlag.Data.SearchFolderID);
					szExtendedFlags += strings::formatmessage(IDS_EXTENDEDFLAGDATASFID, szFlags.c_str());
					break;
				case EFPB_SFTAG:
					szExtendedFlags += strings::formatmessage(IDS_EXTENDEDFLAGDATASFTAG,
						extendedFlag.Data.SearchFolderTag);
					break;
				case EFPB_TODO_VERSION:
					szExtendedFlags += strings::formatmessage(IDS_EXTENDEDFLAGDATATODOVERSION, extendedFlag.Data.ToDoFolderVersion);
					break;
				}

				if (extendedFlag.lpUnknownData.size())
				{
					szExtendedFlags += strings::loadstring(IDS_EXTENDEDFLAGUNKNOWN);
					szExtendedFlags += strings::BinToHexString(extendedFlag.lpUnknownData, true);
				}
			}
		}

		return szExtendedFlags;
	}
}