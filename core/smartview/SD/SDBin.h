#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockT.h>
#include <core/smartview/block/blockBytes.h>
#include <core/interpret/sid.h>
#include <core/smartview/SD/SidBin.h>
#include <core/smartview/SD/ACLBin.h>

namespace smartview
{
	// [MS-DTYP] 2.4.6 SECURITY_DESCRIPTOR
	// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/7d4dac05-9cef-4563-a058-f108abecce1d
	class SDBin : public block
	{
	public:
		SDBin(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB);

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<BYTE>> Revision = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> Sbz1 = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> Control = emptyT<WORD>();
		std::shared_ptr<blockT<DWORD>> OffsetOwner = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> OffsetGroup = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> OffsetSacl = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> OffsetDacl = emptyT<DWORD>();
		std::shared_ptr<SIDBin> OwnerSid;
		std::shared_ptr<SIDBin> GroupSid;
		std::shared_ptr<ACLBin> Sacl;
		std::shared_ptr<ACLBin> Dacl;

		sid::aceType acetype{sid::aceType::Message};
	};
} // namespace smartview