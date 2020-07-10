#include <core/stdafx.h>
#include <core/smartview/encodeEntryID.h>

namespace smartview
{
	_Check_return_ std::wstring EncodeID(ULONG cbEID, _In_ const BYTE* pbSrc)
	{
		std::wstring wzIDEncoded;
		wzIDEncoded.reserve(cbEID);

		// pbSrc is the item Entry ID or the attachment ID
		// cbEID is the size in bytes of pbSrc
		for (ULONG i = 0; i < cbEID; i++, pbSrc++)
		{
			wzIDEncoded.push_back(static_cast<WCHAR>(*pbSrc + kwBaseOffset));
		}

		// wzIDEncoded now contains the entry ID encoded.
		return wzIDEncoded;
	}

	void encodeEntryID::parse() { entryID = blockBytes::parse(parser, parser->getSize()); }

	void encodeEntryID::parseBlocks()
	{
		setText(L"Encoded Entry ID");
		addChild(entryID, EncodeID(entryID->size(), entryID->data()));
	}
} // namespace smartview