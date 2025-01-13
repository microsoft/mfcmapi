#include <core/stdafx.h>
#include <core/smartview/SD/NTSD.h>
#include <core/smartview/SD/SIDBin.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/flags.h>
#include <core/interpret/sid.h>
#include <core/mapi/mapiFunctions.h>

namespace smartview
{
	void NTSD::parse()
	{
		// Grab a parser at the start to pass to NT functions
		const auto sdRoot = parser->getOffset();
		m_SDbin = blockBytes::parse(parser, parser->getSize());
		parser->setOffset(sdRoot);

		Revision = blockT<BYTE>::parse(parser);
		Sbz1 = blockT<BYTE>::parse(parser);
		Control = blockT<WORD>::parse(parser);
		OffsetOwner = blockT<DWORD>::parse(parser);
		auto newOffset = OffsetOwner->getData();
		if (newOffset && newOffset < m_SDbin->size())
		{
			const auto originalOffset = parser->getOffset();
			parser->setOffset(newOffset);
			block::parse<SIDBin>(parser, false);
			parser->setOffset(originalOffset);
		}

		OffsetGroup = blockT<DWORD>::parse(parser);
		newOffset = OffsetGroup->getData();
		if (newOffset && newOffset < m_SDbin->size())
		{
			const auto originalOffset = parser->getOffset();
			parser->setOffset(newOffset);
			block::parse<SIDBin>(parser, false);
			parser->setOffset(originalOffset);
		}

		OffsetSacl = blockT<DWORD>::parse(parser);
		newOffset = OffsetSacl->getData();
		if (newOffset && newOffset < m_SDbin->size())
		{
			const auto originalOffset = parser->getOffset();
			parser->setOffset(newOffset);
			Sacl = block::parse<ACLBin>(parser, false);
			parser->setOffset(originalOffset);
		}

		OffsetDacl = blockT<DWORD>::parse(parser);
		newOffset = OffsetDacl->getData();
		if (newOffset && newOffset < m_SDbin->size())
		{
			const auto originalOffset = parser->getOffset();
			parser->setOffset(newOffset);
			Dacl = block::parse<ACLBin>(parser, false);
			parser->setOffset(originalOffset);
		}
	}

	void NTSD::parseBlocks()
	{
		setText(L"PR_NT_SECURITY_DESCRIPTOR");

		addChild(Revision, L"Revision: 0x%1!02X!", Revision->getData());
		addChild(Sbz1, L"Sbz1: 0x%1!02X!", Sbz1->getData());
		addChild(Control, L"Control: 0x%1!04X!", Control->getData());
		addChild(OffsetOwner, L"OffsetOwner: 0x%1!08X!", OffsetOwner->getData());
		addChild(OffsetGroup, L"OffsetGroup: 0x%1!08X!", OffsetGroup->getData());
		addChild(OffsetSacl, L"OffsetSacl: 0x%1!08X!", OffsetSacl->getData());
		if (Sacl) addChild(Sacl);
		addChild(OffsetDacl, L"OffsetDacl: 0x%1!08X!", OffsetDacl->getData());
		if (Dacl) addChild(Dacl);

		if (m_SDbin)
		{
			// TODO: more accurately break this parsing into blocks with proper offsets
			const auto sd = SDToString(*m_SDbin, acetype);
			auto si = create(L"Security Info");
			addChild(si);
			if (!sd.info.empty())
			{
				si->addChild(m_SDbin, sd.info);
			}

			if (m_SDbin->size() >= 2 * sizeof(WORD))
			{
				const auto sdVersion = SECURITY_DESCRIPTOR_VERSION(m_SDbin->data());
				auto szFlags = flags::InterpretFlags(flagSecurityVersion, sdVersion);
				addHeader(L"Security Version: 0x%1!04X! = %2!ws!", sdVersion, szFlags.c_str());
			}

			addHeader(L"Descriptor");
			addHeader(sd.dacl);
		}
	}
} // namespace smartview