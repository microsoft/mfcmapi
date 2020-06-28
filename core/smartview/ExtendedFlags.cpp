#include <core/stdafx.h>
#include <core/smartview/ExtendedFlags.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void ExtendedFlag::parse()
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
				ExtendedFlags = blockT<DWORD>::parse(parser);
			else
				bBadData = true;
			break;
		case EFPB_CLSIDID:
			if (*Cb == sizeof(GUID))
				SearchFolderID = blockT<GUID>::parse(parser);
			else
				bBadData = true;
			break;
		case EFPB_SFTAG:
			if (*Cb == sizeof(DWORD))
				SearchFolderTag = blockT<DWORD>::parse(parser);
			else
				bBadData = true;
			break;
		case EFPB_TODO_VERSION:
			if (*Cb == sizeof(DWORD))
				ToDoFolderVersion = blockT<DWORD>::parse(parser);
			else
				bBadData = true;
			break;
		default:
			lpUnknownData = blockBytes::parse(parser, *Cb, _MaxBytes);
			break;
		}
	}

	void ExtendedFlag::parseBlocks()
	{
		auto szFlags = flags::InterpretFlags(flagExtendedFolderFlagType, *Id);
		addChild(Id, L"Id = 0x%1!02X! = %2!ws!", Id->getData(), szFlags.c_str());
		Id->addChild(Cb, L"Cb = 0x%1!02X! = %1!d!", Cb->getData());

		switch (*Id)
		{
		case EFPB_FLAGS:
			Id->addChild(
				ExtendedFlags,
				L"Extended Flags = 0x%1!08X! = %2!ws!",
				ExtendedFlags->getData(),
				flags::InterpretFlags(flagExtendedFolderFlag, *ExtendedFlags).c_str());
			break;
		case EFPB_CLSIDID:
			Id->addChild(SearchFolderID, L"SearchFolderID = %1!ws!", guid::GUIDToString(*SearchFolderID).c_str());
			break;
		case EFPB_SFTAG:
			Id->addChild(SearchFolderTag, L"SearchFolderTag = 0x%1!08X!", SearchFolderTag->getData());
			break;
		case EFPB_TODO_VERSION:
			Id->addChild(ToDoFolderVersion, L"ToDoFolderVersion = 0x%1!08X!", ToDoFolderVersion->getData());
			break;
		}

		if (lpUnknownData->size())
		{
			addLabeledChild(L"Unknown Data =", lpUnknownData);
		}
	}

	void ExtendedFlags::parse()
	{
		ULONG ulNumFlags{};
		// Run through the parser once to count the number of flag structs
		for (;;)
		{
			// Must have at least 2 bytes left to have another flag
			if (parser->getSize() < 2) break;
			parser->advance(sizeof BYTE);
			const auto cbData = blockT<BYTE>::parse(parser);
			// Must have at least cbData bytes left to be a valid flag
			if (parser->getSize() < *cbData) break;

			parser->advance(*cbData);
			ulNumFlags++;
		}

		// Now we parse for real
		parser->rewind();

		if (ulNumFlags && ulNumFlags < _MaxEntriesSmall)
		{
			m_pefExtendedFlags.reserve(ulNumFlags);
			for (ULONG i = 0; i < ulNumFlags; i++)
			{
				const auto flag = block::parse<ExtendedFlag>(parser, false);
				m_pefExtendedFlags.push_back(flag);
				if (flag->bBadData) break;
			}
		}
	}

	void ExtendedFlags::parseBlocks()
	{
		setText(L"Extended Flags");
		addSubHeader(L"Number of flags = %1!d!", m_pefExtendedFlags.size());

		if (m_pefExtendedFlags.size())
		{
			int i = 0;
			for (const auto& extendedFlag : m_pefExtendedFlags)
			{
				addChild(extendedFlag, L"Flag[%1!d!]", i++);
			}
		}
	}
} // namespace smartview