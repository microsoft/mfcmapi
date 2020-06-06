#include <core/stdafx.h>
#include <core/smartview/ExtendedFlags.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/block/scratchBlock.h>

namespace smartview
{
	ExtendedFlag::ExtendedFlag(const std::shared_ptr<binaryParser>& parser)
	{
		Id = blockT<BYTE>::parse(parser);
		Cb = blockT<BYTE>::parse(parser);

		// If the structure says there's more bytes than remaining buffer, we're done parsing.
		if (parser->getSize() < *Cb)
		{
			bBadData = true;
			return;
		}

		switch (*Id)
		{
		case EFPB_FLAGS:
			if (*Cb == sizeof(DWORD))
				Data.ExtendedFlags = blockT<DWORD>::parse(parser);
			else
				bBadData = true;
			break;
		case EFPB_CLSIDID:
			if (*Cb == sizeof(GUID))
				Data.SearchFolderID = blockT<GUID>::parse(parser);
			else
				bBadData = true;
			break;
		case EFPB_SFTAG:
			if (*Cb == sizeof(DWORD))
				Data.SearchFolderTag = blockT<DWORD>::parse(parser);
			else
				bBadData = true;
			break;
		case EFPB_TODO_VERSION:
			if (*Cb == sizeof(DWORD))
				Data.ToDoFolderVersion = blockT<DWORD>::parse(parser);
			else
				bBadData = true;
			break;
		default:
			lpUnknownData = blockBytes::parse(parser, *Cb, _MaxBytes);
			break;
		}
	}

	void ExtendedFlags::parse()
	{
		ULONG ulNumFlags{};
		// Run through the parser once to count the number of flag structs
		for (;;)
		{
			// Must have at least 2 bytes left to have another flag
			if (m_Parser->getSize() < 2) break;
			m_Parser->advance(sizeof BYTE);
			const auto cbData = blockT<BYTE>::parse(m_Parser);
			// Must have at least cbData bytes left to be a valid flag
			if (m_Parser->getSize() < *cbData) break;

			m_Parser->advance(*cbData);
			ulNumFlags++;
		}

		// Now we parse for real
		m_Parser->rewind();

		if (ulNumFlags && ulNumFlags < _MaxEntriesSmall)
		{
			m_pefExtendedFlags.reserve(ulNumFlags);
			for (ULONG i = 0; i < ulNumFlags; i++)
			{
				const auto flag = std::make_shared<ExtendedFlag>(m_Parser);
				m_pefExtendedFlags.push_back(flag);
				if (flag->bBadData) break;
			}
		}
	}

	void ExtendedFlags::parseBlocks()
	{
		setText(L"Extended Flags:\r\n");
		addChild(header(L"Number of flags = %1!d!", m_pefExtendedFlags.size()));

		if (m_pefExtendedFlags.size())
		{
			for (const auto& extendedFlag : m_pefExtendedFlags)
			{
				terminateBlock();
				auto szFlags = flags::InterpretFlags(flagExtendedFolderFlagType, *extendedFlag->Id);
				addChild(
					extendedFlag->Id, L"Id = 0x%1!02X! = %2!ws!\r\n", extendedFlag->Id->getData(), szFlags.c_str());
				addChild(extendedFlag->Cb, L"Cb = 0x%1!02X! = %1!d!\r\n", extendedFlag->Cb->getData());

				switch (*extendedFlag->Id)
				{
				case EFPB_FLAGS:
					terminateBlock();
					addChild(
						extendedFlag->Data.ExtendedFlags,
						L"\tExtended Flags = 0x%1!08X! = %2!ws!",
						extendedFlag->Data.ExtendedFlags->getData(),
						flags::InterpretFlags(flagExtendedFolderFlag, *extendedFlag->Data.ExtendedFlags).c_str());
					break;
				case EFPB_CLSIDID:
					terminateBlock();
					addChild(
						extendedFlag->Data.SearchFolderID,
						L"\tSearchFolderID = %1!ws!",
						guid::GUIDToString(*extendedFlag->Data.SearchFolderID).c_str());
					break;
				case EFPB_SFTAG:
					terminateBlock();
					addChild(
						extendedFlag->Data.SearchFolderTag,
						L"\tSearchFolderTag = 0x%1!08X!",
						extendedFlag->Data.SearchFolderTag->getData());
					break;
				case EFPB_TODO_VERSION:
					terminateBlock();
					addChild(
						extendedFlag->Data.ToDoFolderVersion,
						L"\tToDoFolderVersion = 0x%1!08X!",
						extendedFlag->Data.ToDoFolderVersion->getData());
					break;
				}

				if (extendedFlag->lpUnknownData->size())
				{
					terminateBlock();
					addChild(labeledBlock(L"\tUnknown Data = ", extendedFlag->lpUnknownData));
				}
			}
		}
	}
} // namespace smartview