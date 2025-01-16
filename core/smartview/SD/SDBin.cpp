#include <core/stdafx.h>
#include <core/smartview/SD/SDBin.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/flags.h>
#include <core/interpret/sid.h>
#include <core/mapi/mapiFunctions.h>
#include <core/smartview/SD/SidBin.h>

namespace smartview
{
	SDBin::SDBin(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB)
	{
		switch (mapi::GetMAPIObjectType(lpMAPIProp))
		{
		case MAPI_STORE:
		case MAPI_ADDRBOOK:
		case MAPI_FOLDER:
		case MAPI_ABCONT:
			acetype = sid::aceType::Container;
			break;
		}

		if (bFB) acetype = sid::aceType::FreeBusy;
	}

	void SDBin::parse()
	{
		const auto baseOffset = parser->getOffset();
		auto originalOffset = size_t{};
		const auto sdSize = parser->getSize();

		Revision = blockT<BYTE>::parse(parser);
		Sbz1 = blockT<BYTE>::parse(parser);
		Control = blockT<WORD>::parse(parser);
		OffsetOwner = blockT<DWORD>::parse(parser);

		// Read from offsets now - first remember where we are
		// We'll consider anything after our last read to be junk
		auto postSdOffset = parser->getOffset();
		auto newOffset = *OffsetOwner + baseOffset;
		if (*OffsetOwner && newOffset < sdSize)
		{
			originalOffset = parser->getOffset();
			parser->setOffset(newOffset);
			OwnerSid = block::parse<SIDBin>(parser, false);
			postSdOffset = max(postSdOffset, parser->getOffset());
			parser->setOffset(originalOffset);
		}

		OffsetGroup = blockT<DWORD>::parse(parser);
		newOffset = *OffsetGroup + baseOffset;
		if (*OffsetGroup && newOffset < sdSize)
		{
			originalOffset = parser->getOffset();
			parser->setOffset(newOffset);
			GroupSid = block::parse<SIDBin>(parser, false);
			postSdOffset = max(postSdOffset, parser->getOffset());
			parser->setOffset(originalOffset);
		}

		OffsetSacl = blockT<DWORD>::parse(parser);
		newOffset = *OffsetSacl + baseOffset;
		if (*OffsetSacl && newOffset < sdSize)
		{
			originalOffset = parser->getOffset();
			parser->setOffset(newOffset);
			Sacl = block::parse<ACLBin>(parser, false);
			postSdOffset = max(postSdOffset, parser->getOffset());
			parser->setOffset(originalOffset);
		}

		OffsetDacl = blockT<DWORD>::parse(parser);
		newOffset = *OffsetDacl + baseOffset;
		if (*OffsetDacl && newOffset < sdSize)
		{
			originalOffset = parser->getOffset();
			parser->setOffset(newOffset);
			Dacl = block::parse<ACLBin>(parser, false);
			postSdOffset = max(postSdOffset, parser->getOffset());
			parser->setOffset(originalOffset);
		}

		// Having read everything, set our offset to the end of the SD
		parser->setOffset(postSdOffset);
	}

	void SDBin::parseBlocks()
	{
		setText(L"Security Descriptor");

		addChild(Revision, L"Revision: 0x%1!02X!", Revision->getData());
		addChild(Sbz1, L"Sbz1: 0x%1!02X!", Sbz1->getData());
		addChild(Control, L"Control: 0x%1!04X!", Control->getData());
		addChild(OffsetOwner, L"OffsetOwner: 0x%1!08X!", OffsetOwner->getData());
		addChild(OffsetGroup, L"OffsetGroup: 0x%1!08X!", OffsetGroup->getData());
		addChild(OffsetSacl, L"OffsetSacl: 0x%1!08X!", OffsetSacl->getData());
		addChild(OffsetDacl, L"OffsetDacl: 0x%1!08X!", OffsetDacl->getData());
		if (OwnerSid) addLabeledChild(L"OwnerSid", OwnerSid);
		if (GroupSid) addLabeledChild(L"GroupSid", GroupSid);
		if (Sacl) addLabeledChild(L"Sacl", Sacl);
		if (Dacl) addLabeledChild(L"Dacl", Dacl);
	}
} // namespace smartview