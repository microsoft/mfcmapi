#include <core/stdafx.h>
#include <core/mapi/mapiBinaryFunctions.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/columnTags.h>
#include <core/interpret/guid.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/utility/error.h>

namespace mapi::binary
{
	bool ValidEntryID(_In_reads_bytes_(size) const BYTE* bytes, size_t size, unsigned depth = 0)
	{
		if (!bytes || size < 20 || depth >= 25) return false;

		GUID provider{};
		if (memcpy_s(&provider, sizeof(provider), bytes + 4, sizeof(provider))) return false;

		if (provider == guid::WAB_GUID) return size > 21 && ValidEntryID(bytes + 21, size - 21, depth + 1);

		if (provider == guid::muidContabDLL)
		{
			if (size < 36) return false;

			ULONG type = 0;
			ULONG embeddedSize = 0;
			if (memcpy_s(&type, sizeof(type), bytes + 24, sizeof(type)) ||
				memcpy_s(&embeddedSize, sizeof(embeddedSize), bytes + 32, sizeof(embeddedSize)))
			{
				return false;
			}

			return (type == CONTAB_USER || type == CONTAB_DISTLIST) && embeddedSize <= size - 36 &&
				   ValidEntryID(bytes + 36, embeddedSize, depth + 1);
		}

		if (provider == guid::muidOOP) return size >= 24 && bytes[20] == 0 && bytes[21] == 0;

		return true;
	}

	BinaryTableRow DecodeEntryID(const SBinary& entryID)
	{
		BinaryTableRow member;
		if (!entryID.cb || !entryID.lpb)
		{
			member.status = WC_H(MAPI_E_INVALID_ENTRYID);
			return member;
		}

		member.entryID.assign(entryID.lpb, entryID.lpb + entryID.cb);
		if (!ValidEntryID(entryID.lpb, entryID.cb))
		{
			member.status = WC_H(MAPI_E_INVALID_ENTRYID);
			return member;
		}

		member.openEntryID = member.entryID;
		auto parser = std::make_shared<smartview::binaryParser>(entryID.cb, entryID.lpb);
		member.decoded = smartview::block::parse<smartview::EntryIdStruct>(parser, true)->toString();

		// WAB member wrappers have a 20-byte header followed by a one-byte member type.
		// Keep the original bytes for inspection, but pass the embedded EntryID to OpenEntry.
		for (unsigned depth = 0; depth < 25; ++depth)
		{
			const auto& bytes = member.openEntryID;
			if (bytes.size() < 20)
			{
				member.status = WC_H(MAPI_E_INVALID_ENTRYID);
				return member;
			}

			GUID provider{};
			if (memcpy_s(&provider, sizeof(provider), bytes.data() + 4, sizeof(provider)))
			{
				member.status = WC_H(MAPI_E_INVALID_ENTRYID);
				return member;
			}

			if (provider == guid::WAB_GUID)
			{
				if (bytes.size() <= 21)
				{
					member.status = WC_H(MAPI_E_INVALID_ENTRYID);
					return member;
				}

				member.openEntryID = std::vector<BYTE>(bytes.begin() + 21, bytes.end());
				continue;
			}

			if (provider == guid::muidOOP)
			{
				auto oneOffParser = std::make_shared<smartview::binaryParser>(bytes.size() - 20, bytes.data() + 20);
				const auto oneOff = smartview::block::parse<smartview::OneOffRecipientObject>(oneOffParser, false);
				if (!oneOff->getRecipient(member.name, member.addressType, member.address))
				{
					member.status = WC_H(MAPI_E_INVALID_ENTRYID);
				}
			}

			return member;
		}

		member.status = WC_H(MAPI_E_TOO_COMPLEX);

		return member;
	}

	_Check_return_ HRESULT ExtractBinaryTableRows(
		const SPropValue& property,
		ULONG expectedTag,
		std::vector<BinaryTableRow>& members,
		bool entryIDs)
	{
		members.clear();

		if (PROP_TYPE(property.ulPropTag) == PT_ERROR)
		{
			return WC_H(FAILED(property.Value.err) ? property.Value.err : MAPI_E_INVALID_TYPE);
		}

		if (PROP_TYPE(expectedTag) != PT_MV_BINARY || property.ulPropTag != expectedTag)
			return WC_H(MAPI_E_INVALID_TYPE);

		const auto& array = property.Value.MVbin;
		if (array.cValues && !array.lpbin) return WC_H(MAPI_E_INVALID_PARAMETER);
		if (array.cValues > smartview::_MaxEntriesEnormous) return WC_H(MAPI_E_TOO_COMPLEX);

		members.reserve(array.cValues);
		for (ULONG i = 0; i < array.cValues; ++i)
		{
			if (entryIDs)
			{
				members.emplace_back(DecodeEntryID(array.lpbin[i]));
			}
			else
			{
				const auto& binary = array.lpbin[i];
				if (binary.cb && !binary.lpb)
				{
					members.clear();
					return WC_H(MAPI_E_INVALID_PARAMETER);
				}

				BinaryTableRow member;
				if (binary.cb) member.entryID.assign(binary.lpb, binary.lpb + binary.cb);
				members.emplace_back(std::move(member));
			}
		}

		return S_OK;
	}

	_Check_return_ HRESULT
	ReadBinaryTableRows(_In_ LPMAPIPROP message, ULONG propTag, std::vector<BinaryTableRow>& members, bool entryIDs)
	{
		members.clear();

		if (!message) return WC_H(MAPI_E_INVALID_PARAMETER);
		if (PROP_TYPE(propTag) != PT_MV_BINARY) return WC_H(MAPI_E_INVALID_TYPE);

		SizedSPropTagArray(1, tags) = {1, {propTag}};
		ULONG count = 0;
		LPSPropValue property = nullptr;
		auto hRes = WC_H_GETPROPS(message->GetProps(reinterpret_cast<LPSPropTagArray>(&tags), 0, &count, &property));
		if (SUCCEEDED(hRes))
		{
			hRes = count == 1 && property ? ExtractBinaryTableRows(*property, propTag, members, entryIDs)
										  : WC_H(MAPI_E_NOT_FOUND);
		}

		MAPIFreeBuffer(property);

		return hRes;
	}

	_Check_return_ HRESULT
	CreateBinaryTable(const std::vector<BinaryTableRow>& members, _Deref_out_ LPMAPITABLE* table, bool entryIDs)
	{
		if (!table) return WC_H(MAPI_E_INVALID_PARAMETER);
		*table = nullptr;

		if (members.size() > ULONG_MAX) return WC_H(MAPI_E_TOO_BIG);

		LPTABLEDATA data = nullptr;
		auto hRes = WC_H(CreateTable(
			&IID_IMAPITableData,
			MAPIAllocateBuffer,
			MAPIAllocateMore,
			MAPIFreeBuffer,
			nullptr,
			TBLTYPE_SNAPSHOT,
			PR_ROWID,
			entryIDs ? &columns::sptEntryIDCols.tags : &columns::sptBinaryCols.tags,
			&data));
		if (SUCCEEDED(hRes) && !data) hRes = WC_H(E_UNEXPECTED);

		for (size_t i = 0; SUCCEEDED(hRes) && i < members.size(); ++i)
		{
			const auto& member = members[i];
			if (member.entryID.size() > LONG_MAX)
			{
				hRes = WC_H(MAPI_E_TOO_BIG);
				break;
			}

			ULONG rowID = static_cast<ULONG>(i + 1);
			SPropValue props[columns::ENTRYIDTAGS_NUM_COLS]{};
			for (ULONG col = 0; col < _countof(props); ++col)
			{
				props[col].ulPropTag = columns::sptEntryIDCols.tags.aulPropTag[col];
			}

			props[columns::eidtagPR_INSTANCE_KEY].Value.bin = {sizeof(rowID), reinterpret_cast<LPBYTE>(&rowID)};
			props[columns::eidtagPR_ENTRYID].Value.bin = {
				static_cast<ULONG>(member.entryID.size()), const_cast<LPBYTE>(member.entryID.data())};
			props[columns::eidtagPR_ROWID].Value.l = rowID;
			props[columns::eidtagPR_DISPLAY_NAME].Value.lpszW = const_cast<LPWSTR>(member.name.c_str());
			props[columns::eidtagPR_ADDRTYPE].Value.lpszW = const_cast<LPWSTR>(member.addressType.c_str());
			props[columns::eidtagPR_EMAIL_ADDRESS].Value.lpszW = const_cast<LPWSTR>(member.address.c_str());
			props[columns::eidtagPR_STATUS_CODE].Value.l = member.status;
			props[columns::eidtagPR_BODY].Value.lpszW = const_cast<LPWSTR>(member.decoded.c_str());
			SRow row{0, _countof(props), props};
			if (!entryIDs)
			{
				props[1].Value.bin = {};
				props[3].ulPropTag = PR_ATTACH_SIZE;
				props[3].Value.l = static_cast<LONG>(member.entryID.size());
				props[4].ulPropTag = PR_ATTACH_DATA_BIN;
				props[4].Value.bin = {
					static_cast<ULONG>(member.entryID.size()), const_cast<LPBYTE>(member.entryID.data())};
				row.cValues = columns::sptBinaryCols.tags.cValues;
			}

			hRes = WC_H(data->HrModifyRow(&row));
		}

		if (SUCCEEDED(hRes)) hRes = WC_H(data->HrGetView(nullptr, nullptr, 0, table));
		if (data) data->Release();

		return hRes;
	}

	_Check_return_ LPMAPIPROP OpenEntryID(LPADRBOOK addressBook, LPMAPISESSION session, BinaryTableRow& member)
	{
		if (!ValidEntryID(member.openEntryID.data(), member.openEntryID.size()) ||
			member.status == MAPI_E_INVALID_ENTRYID)
		{
			member.status = WC_H(MAPI_E_INVALID_ENTRYID);
			return nullptr;
		}

		SBinary entry{static_cast<ULONG>(member.openEntryID.size()), member.openEntryID.data()};
		LPMAPIPROP object = nullptr;
		member.status = MAPI_E_NO_SUPPORT;
		if (addressBook)
		{
			ULONG type = 0;
			LPUNKNOWN unknown = nullptr;
			member.status = WC_H(
				addressBook->OpenEntry(entry.cb, reinterpret_cast<LPENTRYID>(entry.lpb), nullptr, 0, &type, &unknown));
			if (SUCCEEDED(member.status) && unknown) object = safe_cast<LPMAPIPROP>(unknown);
			if (unknown) unknown->Release();
			if (SUCCEEDED(member.status) && !object) member.status = WC_H(E_NOINTERFACE);
		}

		if (!object && session)
		{
			GUID provider{};
			if (memcpy_s(&provider, sizeof(provider), entry.lpb + 4, sizeof(provider)))
			{
				member.status = WC_H(MAPI_E_INVALID_ENTRYID);
				return nullptr;
			}

			if (provider == guid::muidContabDLL && entry.cb >= 36)
			{
				ULONG embeddedSize = 0;
				if (memcpy_s(&embeddedSize, sizeof(embeddedSize), entry.lpb + 32, sizeof(embeddedSize)))
				{
					member.status = WC_H(MAPI_E_INVALID_ENTRYID);
					return nullptr;
				}

				ULONG size = 0;
				LPBYTE bytes = nullptr;
				// The legacy helper returns a borrowed pointer; validate its advertised size first.
				if (embeddedSize && embeddedSize <= entry.cb - 36 &&
					UnwrapContactEntryID(entry.cb, entry.lpb, &size, &bytes))
				{
					entry = {size, bytes};
				}
			}

			ULONG type = 0;
			LPUNKNOWN unknown = nullptr;
			member.status =
				WC_H(session->OpenEntry(entry.cb, reinterpret_cast<LPENTRYID>(entry.lpb), nullptr, 0, &type, &unknown));
			if (SUCCEEDED(member.status) && unknown) object = safe_cast<LPMAPIPROP>(unknown);
			if (unknown) unknown->Release();
			if (SUCCEEDED(member.status) && !object) member.status = WC_H(E_NOINTERFACE);
		}

		if (object) member.status = S_OK;

		return object;
	}

} // namespace mapi::binary
