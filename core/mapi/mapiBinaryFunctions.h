#pragma once
#include <core/mapi/extraPropTags.h>

namespace mapi::binary
{
	struct BinaryTableRow
	{
		std::vector<BYTE> entryID;
		std::vector<BYTE> openEntryID;
		std::wstring name;
		std::wstring addressType;
		std::wstring address;
		std::wstring decoded;
		HRESULT status = S_OK;
	};

	// These helpers never create named properties, expand groups, or modify the message.
	_Check_return_ HRESULT
	ReadBinaryTableRows(_In_ LPMAPIPROP message, ULONG propTag, std::vector<BinaryTableRow>& members, bool entryIDs);
	_Check_return_ HRESULT
	CreateBinaryTable(const std::vector<BinaryTableRow>& members, _Deref_out_ LPMAPITABLE* table, bool entryIDs);
	_Check_return_ HRESULT ExtractBinaryTableRows(
		const SPropValue& property,
		ULONG expectedTag,
		std::vector<BinaryTableRow>& members,
		bool entryIDs);
	BinaryTableRow DecodeEntryID(const SBinary& entryID);
	_Check_return_ LPMAPIPROP OpenEntryID(LPADRBOOK addressBook, LPMAPISESSION session, BinaryTableRow& member);

} // namespace mapi::binary
