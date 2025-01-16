#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/SD/NTSD.h>
#include <core/smartview/SD/SDBin.h>

namespace smartview
{
	class NamedProp : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<WORD>> tag = emptyT<WORD>();
		std::shared_ptr<blockT<GUID>> guid = emptyT<GUID>();
		std::shared_ptr<blockT<BYTE>> kind = emptyT<BYTE>();
		std::shared_ptr<blockT<DWORD>> id = emptyT<DWORD>();
		std::shared_ptr<blockStringW> name = emptySW();
	};

	// PR_NT_SECURITY_DESCRIPTOR
	// https://github.com/microsoft/MAPIStubLibrary/blob/main/include/EdkMdb.h
	//
	//	Transfer version for PR_NT_SECURITY_DESCRIPTOR.
	//
	//	When retrieving the security descriptor for an object, the SD returned is
	//	actually composed of the following structure:
	//
	//		2 BYTES					Padding data length (including version)
	//		2 BYTES					Version
	//		4 BYTES					Security Information (for SetPrivateObjectSecurity)
	//		<0 or more>
	//			2 BYTES					Property Tag
	//			16 BYTES				Named Property GUID
	//			1 BYTE					Named property "kind"
	//			if (kind == MNID_ID)
	//				4 BYTES				Named property ID
	//			else
	//				<null terminated property name in UNICODE!!!!!>
	//		Actual Security Descriptor
	class NTSD : public block
	{
	public:
		NTSD(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB);
		NTSD(_In_ sid::aceType _acetype) : acetype(_acetype){};

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<WORD>> Padding = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> Version = emptyT<WORD>();
		std::shared_ptr<blockT<DWORD>> SecurityInformation = emptyT<DWORD>();
		std::vector<std::shared_ptr<NamedProp>> NamedProperties;
		std::shared_ptr<SDBin> SD;

		sid::aceType acetype{sid::aceType::Message};
	};
} // namespace smartview