#include <core/stdafx.h>
#include <core/smartview/decodeEntryID.h>
#include <core/smartview/encodeEntryID.h>

namespace smartview
{
	_Check_return_ std::wstring DecodeID(size_t cbBuffer, _In_count_(cbBuffer) const BYTE* lpbBuffer)
	{
		if (cbBuffer % 2) return strings::emptystring;

		const auto cbDecodedBuffer = cbBuffer / 2;
		auto lpDecoded = std::vector<BYTE>();
		lpDecoded.reserve(cbDecodedBuffer);

		// Subtract kwBaseOffset from every character and place result in lpDecoded
		auto lpwzSrc = reinterpret_cast<LPCWSTR>(lpbBuffer);
		for (size_t i = 0; i < cbDecodedBuffer; i++, lpwzSrc++)
		{
			lpDecoded.push_back(static_cast<BYTE>(*lpwzSrc - kwBaseOffset));
		}

		return strings::BinToHexString(lpDecoded, true);
	}

	void decodeEntryID::parse() { entryID = blockBytes::parse(parser, parser->getSize()); }

	void decodeEntryID::parseBlocks()
	{
		setText(L"Decoded Entry ID");
		addChild(entryID, DecodeID(entryID->size(), entryID->data()));
	}
} // namespace smartview