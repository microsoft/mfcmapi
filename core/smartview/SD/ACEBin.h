#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>
#include <core/interpret/sid.h>
#include <core/smartview/SD/SIDBin.h>

namespace smartview
{
	// [MS-DTYP] 2.4.4 ACE
	// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/d06e5a81-176e-46c6-9cf7-9137aad4455e
	// [MS-DTYP] 2.4.4.1 ACE_HEADER
	// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/628ebb1d-c509-4ea0-a10f-77ef97ca4586
	// [MS-DTYP] 2.4.4.2 ACCESS_ALLOWED_ACE
	// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/72e7c7ea-bc02-4c74-a619-818a16bf6adb
	// [MS-DTYP] 2.4.4.3 ACCESS_ALLOWED_OBJECT_ACE
	// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/c79a383c-2b3f-4655-abe7-dcbb7ce0cfbe
	// [MS-DTYP] 2.4.4.4 ACCESS_DENIED_ACE
	// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/b1e1321d-5816-4513-be67-b65d8ae52fe8
	// [MS-DTYP] 2.4.4.5 ACCESS_DENIED_OBJECT_ACE
	// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/8720fcf3-865c-4557-97b1-0b3489a6c270

	class ACEBin : public block
	{
	public:
		ACEBin(sid::aceType acetype);

	private:
		sid::aceType acetype{sid::aceType::Message};
		std::shared_ptr<blockT<BYTE>> AceType = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> AceFlags = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> AceSize = emptyT<WORD>();
		std::shared_ptr<blockT<DWORD>> Mask = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> Flags = emptyT<DWORD>();
		std::shared_ptr<blockT<GUID>> ObjectType = emptyT<GUID>();
		std::shared_ptr<blockT<GUID>> InheritedObjectType = emptyT<GUID>();
		std::shared_ptr<SIDBin> SidStart;

		void parse() override;
		void parseBlocks() override;
	};
} // namespace smartview