#include <core/stdafx.h>
#include <core/smartview/SD/ACLBin.h>

namespace smartview
{
	void ACLBin::parse()
	{
		Revision = blockT<BYTE>::parse(parser);
		Sbz1 = blockT<BYTE>::parse(parser);
		AclSize = blockT<WORD>::parse(parser);
		AceCount = blockT<WORD>::parse(parser);
		Sbz2 = blockT<WORD>::parse(parser);
		for (auto i = 0; i < AceCount->getData(); i++)
		{
			const auto ace = std::make_shared<ACEBin>(sid::aceType::Message);
			ace->block::parse(parser, false);
			if (!ace->isSet()) break;
			aces.push_back(ace);
		}
	};

	void ACLBin::parseBlocks()
	{
		setText(L"ACL");
		addChild(Revision, L"Revision: 0x%1!02X!", Revision->getData());
		addChild(Sbz1, L"Sbz1: 0x%1!02X!", Sbz1->getData());
		addChild(AclSize, L"AclSize: 0x%1!04X!", AclSize->getData());
		addChild(AceCount, L"AceCount: 0x%1!04X!", AceCount->getData());
		addChild(Sbz2, L"Sbz2: 0x%1!04X!", Sbz2->getData());

		for (const auto& ace : aces)
		{
			addChild(ace);
		}
	};
} // namespace smartview