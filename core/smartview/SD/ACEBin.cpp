#include <core/stdafx.h>
#include <core/smartview/SD/ACEBin.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/guid.h>

namespace smartview
{
	ACEBin::ACEBin(sid::aceType acetype) { this->acetype = acetype; }

	void ACEBin::parse()
	{
		// Header
		AceType = blockT<BYTE>::parse(parser);
		AceFlags = blockT<BYTE>::parse(parser);
		AceSize = blockT<WORD>::parse(parser);

		// Specific ACE types
		switch (AceType->getData())
		{
		case ACCESS_ALLOWED_ACE_TYPE: // ACCESS_ALLOWED_ACE
			Mask = blockT<DWORD>::parse(parser);
			SidStart = block::parse<SIDBin>(parser, false);
			break;
		case ACCESS_DENIED_ACE_TYPE: // ACCESS_DENIED_ACE
			Mask = blockT<DWORD>::parse(parser);
			SidStart = block::parse<SIDBin>(parser, false);
			break;
		case ACCESS_ALLOWED_OBJECT_ACE_TYPE: // ACCESS_ALLOWED_OBJECT_ACE
			Mask = blockT<DWORD>::parse(parser);
			Flags = blockT<DWORD>::parse(parser);
			ObjectType = blockT<GUID>::parse(parser);
			InheritedObjectType = blockT<GUID>::parse(parser);
			SidStart = block::parse<SIDBin>(parser, false);
			break;
		case ACCESS_DENIED_OBJECT_ACE_TYPE: // ACCESS_DENIED_OBJECT_ACE
			Mask = blockT<DWORD>::parse(parser);
			Flags = blockT<DWORD>::parse(parser);
			ObjectType = blockT<GUID>::parse(parser);
			InheritedObjectType = blockT<GUID>::parse(parser);
			SidStart = block::parse<SIDBin>(parser, false);
			break;
		}
	};

	void ACEBin::parseBlocks()
	{
		setText(L"ACE");
		const auto aceType = AceType->getData();
		auto szAceType = flags::InterpretFlags(flagACEType, aceType);
		addChild(AceType, L"Type: 0x%1!02X! = %2!ws!", aceType, szAceType.c_str());
		const auto aceFlags = AceFlags->getData();
		auto szAceFlags = flags::InterpretFlags(flagACEFlag, aceFlags);
		addChild(AceFlags, L"Flags: 0x%1!02X! = %2!ws!", aceFlags, szAceFlags.c_str());
		addChild(AceSize, L"Size: 0x%1!04X!", AceSize->getData());

		auto szAceMask = std::wstring{};
		switch (acetype)
		{
		case sid::aceType::Container:
			szAceMask = flags::InterpretFlags(flagACEMaskContainer, Mask->getData());
			break;
		case sid::aceType::Message:
			szAceMask = flags::InterpretFlags(flagACEMaskNonContainer, Mask->getData());
			break;
		case sid::aceType::FreeBusy:
			szAceMask = flags::InterpretFlags(flagACEMaskFreeBusy, Mask->getData());
			break;
		};
		addChild(Mask, L"Mask: 0x%1!08X! = %2!ws!", Mask->getData(), szAceMask.c_str());

		addChild(Flags, L"Flags: 0x%1!08X!", Flags->getData());

		addChild(ObjectType, L"ObjectType: %1!ws!", guid::GUIDToStringAndName(ObjectType->getData()).c_str());
		addChild(
			InheritedObjectType,
			L"InheritedObjectType: %1!ws!",
			guid::GUIDToStringAndName(InheritedObjectType->getData()).c_str());
		addChild(SidStart);
	};
} // namespace smartview