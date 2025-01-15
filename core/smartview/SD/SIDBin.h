#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-DTYP] 2.4.2.2 SID--Packet Representation
	// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/f992ad60-0fe4-4b87-9fed-beb478836861
	class SIDBin : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<BYTE>> Revision = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> SubAuthorityCount = emptyT<BYTE>();
		std::shared_ptr<blockBytes> IdentifierAuthority = emptyBB(); // 6 bytes
		std::vector<std::shared_ptr<blockT<DWORD>>> SubAuthority;

		// We keep this for a call to LookupAccountSid
		std::shared_ptr<blockBytes> m_SIDbin = emptyBB();
	};
} // namespace smartview