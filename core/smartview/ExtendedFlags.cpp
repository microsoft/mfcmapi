#include <core/stdafx.h>
#include <core/smartview/ExtendedFlags.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void ExtendedFlags::Parse()
	{
		// Run through the parser once to count the number of flag structs
		for (;;)
		{
			// Must have at least 2 bytes left to have another flag
			if (m_Parser.RemainingBytes() < 2) break;
			(void) m_Parser.Get<BYTE>();
			const auto cbData = m_Parser.Get<BYTE>();
			// Must have at least cbData bytes left to be a valid flag
			if (m_Parser.RemainingBytes() < cbData) break;

			m_Parser.advance(cbData);
			m_ulNumFlags++;
		}

		// Now we parse for real
		m_Parser.rewind();

		if (m_ulNumFlags && m_ulNumFlags < _MaxEntriesSmall)
		{
			auto bBadData = false;

			m_pefExtendedFlags.reserve(m_ulNumFlags);
			for (ULONG i = 0; i < m_ulNumFlags; i++)
			{
				ExtendedFlag extendedFlag;

				extendedFlag.Id = m_Parser.Get<BYTE>();
				extendedFlag.Cb = m_Parser.Get<BYTE>();

				// If the structure says there's more bytes than remaining buffer, we're done parsing.
				if (m_Parser.RemainingBytes() < extendedFlag.Cb)
				{
					m_pefExtendedFlags.push_back(extendedFlag);
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
					m_pefExtendedFlags.push_back(extendedFlag);
					m_ulNumFlags = i;
					break;
				}

				m_pefExtendedFlags.push_back(extendedFlag);
			}
		}
	}

	void ExtendedFlags::ParseBlocks()
	{
		setRoot(L"Extended Flags:\r\n");
		addHeader(L"Number of flags = %1!d!", m_ulNumFlags);

		if (m_pefExtendedFlags.size())
		{
			for (const auto& extendedFlag : m_pefExtendedFlags)
			{
				terminateBlock();
				auto szFlags = flags::InterpretFlags(flagExtendedFolderFlagType, extendedFlag.Id);
				addBlock(extendedFlag.Id, L"Id = 0x%1!02X! = %2!ws!\r\n", extendedFlag.Id.getData(), szFlags.c_str());
				addBlock(extendedFlag.Cb, L"Cb = 0x%1!02X! = %1!d!\r\n", extendedFlag.Cb.getData());

				switch (extendedFlag.Id)
				{
				case EFPB_FLAGS:
					terminateBlock();
					addBlock(
						extendedFlag.Data.ExtendedFlags,
						L"\tExtended Flags = 0x%1!08X! = %2!ws!",
						extendedFlag.Data.ExtendedFlags.getData(),
						flags::InterpretFlags(flagExtendedFolderFlag, extendedFlag.Data.ExtendedFlags).c_str());
					break;
				case EFPB_CLSIDID:
					terminateBlock();
					addBlock(
						extendedFlag.Data.SearchFolderID,
						L"\tSearchFolderID = %1!ws!",
						guid::GUIDToString(extendedFlag.Data.SearchFolderID).c_str());
					break;
				case EFPB_SFTAG:
					terminateBlock();
					addBlock(
						extendedFlag.Data.SearchFolderTag,
						L"\tSearchFolderTag = 0x%1!08X!",
						extendedFlag.Data.SearchFolderTag.getData());
					break;
				case EFPB_TODO_VERSION:
					terminateBlock();
					addBlock(
						extendedFlag.Data.ToDoFolderVersion,
						L"\tToDoFolderVersion = 0x%1!08X!",
						extendedFlag.Data.ToDoFolderVersion.getData());
					break;
				}

				if (extendedFlag.lpUnknownData.size())
				{
					terminateBlock();
					addHeader(L"\tUnknown Data = ");
					addBlock(extendedFlag.lpUnknownData);
				}
			}
		}
	}
} // namespace smartview