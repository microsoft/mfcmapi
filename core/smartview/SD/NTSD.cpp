#include <core/stdafx.h>
#include <core/smartview/SD/NTSD.h>
#include <core/smartview/SD/SIDBin.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/flags.h>
#include <core/interpret/sid.h>
#include <core/mapi/mapiFunctions.h>
#include <core/interpret/guid.h>

namespace smartview
{
	void NamedProp::parse()
	{
		tag = blockT<WORD>::parse(parser);
		guid = blockT<GUID>::parse(parser);
		kind = blockT<BYTE>::parse(parser);
		if (*kind == MNID_ID)
			id = blockT<DWORD>::parse(parser);
		else
			name = blockStringW::parse(parser);
	}
	void NamedProp::parseBlocks()
	{
		setText(L"NamedProp");
		addChild(tag, L"Tag = 0x%1!04X!", tag->getData());
		addChild(guid, L"GUID = %1!ws!", guid::GUIDToString(*guid).c_str());
		addChild(kind, L"Kind = 0x%1!02X!", kind->getData());
		if (*kind == MNID_ID)
			addChild(id, L"ID = 0x%1!08X!", id->getData());
		else
			addChild(name, L"Name = %1!ws!", name->c_str());
	}

	NTSD::NTSD(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB)
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

	void NTSD::parse()
	{
		const auto baseOffset = parser->getOffset();
		const auto bufferSize = parser->getSize();
		Padding = blockT<WORD>::parse(parser);
		Version = blockT<WORD>::parse(parser);
		SecurityInformation = blockT<DWORD>::parse(parser);
		const auto bytesConsumed = parser->getOffset() - baseOffset;
		const auto namedPropSize =
			(bufferSize > *Padding && *Padding >= bytesConsumed) ? *Padding - bytesConsumed : parser->getSize();

		if (namedPropSize > 0)
		{
			parser->setCap(namedPropSize);
			while (true)
			{
				const auto np = block::parse<NamedProp>(parser, false);
				if (!np->isSet()) break;
				NamedProperties.push_back(np);
			}

			parser->clearCap();
		}

		if (*Padding < bufferSize)
		{
			parser->setOffset(baseOffset + *Padding);
			SD = std::make_shared<SDBin>(acetype);
			SD->block::parse(parser, false);
		}
	}

	void NTSD::parseBlocks()
	{
		setText(L"PR_NT_SECURITY_DESCRIPTOR");
		addChild(Padding, L"Padding: 0x%1!04X!", Padding->getData());
		addChild(
			Version,
			L"Version: 0x%1!04X! = %2!ws!",
			Version->getData(),
			flags::InterpretFlags(flagSecurityVersion, *Version).c_str());
		addChild(
			SecurityInformation,
			L"Security Information: 0x%1!08X! = %2!ws!",
			SecurityInformation->getData(),
			flags::InterpretFlags(flagSecurityInfo, *SecurityInformation).c_str());

		for (const auto& np : NamedProperties)
		{
			addChild(np);
		}

		addChild(SD);
	}
} // namespace smartview