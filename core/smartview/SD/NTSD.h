#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
//#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>
#include <core/interpret/sid.h>
#include <core/smartview/SD/NTSD.h>
#include <core/smartview/SD/ACLBin.h>

namespace smartview
{
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
	private:
		void parse() override;
		void parseBlocks() override;

		sid::aceType acetype{sid::aceType::Message};
		std::shared_ptr<blockBytes> m_SDbin = emptyBB();

		std::shared_ptr<blockT<BYTE>> Revision = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> Sbz1 = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> Control = emptyT<WORD>();
		std::shared_ptr<blockT<DWORD>> OffsetOwner = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> OffsetGroup = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> OffsetSacl = emptyT<DWORD>();
		std::shared_ptr<ACLBin> Sacl;
		std::shared_ptr<blockT<DWORD>> OffsetDacl = emptyT<DWORD>();
		std::shared_ptr<ACLBin> Dacl;
	};
} // namespace smartview