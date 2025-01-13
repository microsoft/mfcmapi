#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>
#include <core/smartview/SD/ACEBin.h>

namespace smartview
{
	// [MS-DTYP] 2.4.5 ACL
	// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/32d72257-0e7c-4782-bc2a-405af4d5469d
	class ACLBin : public block
	{
	private:
		std::shared_ptr<blockT<BYTE>> Revision = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> Sbz1 = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> AclSize = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> AceCount = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> Sbz2 = emptyT<WORD>();
		std::vector<std::shared_ptr<ACEBin>> ace;

		void parse() override;
		void parseBlocks() override;
	};
} // namespace smartview