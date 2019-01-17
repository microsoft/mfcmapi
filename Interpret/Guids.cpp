#include <StdAfx.h>
#include <AddIns.h>

#include <Interpret/guids.h>
#include <core/utility/strings.h>

namespace guid
{
	std::wstring GUIDToString(_In_ GUID guid) { return GUIDToString(&guid); }

	std::wstring GUIDToString(_In_opt_ LPCGUID lpGUID)
	{
		GUID nullGUID = {};

		if (!lpGUID)
		{
			lpGUID = &nullGUID;
		}

		return strings::format(
			L"{%.8X-%.4X-%.4X-%.2X%.2X-%.2X%.2X%.2X%.2X%.2X%.2X}", // STRING_OK
			lpGUID->Data1,
			lpGUID->Data2,
			lpGUID->Data3,
			lpGUID->Data4[0],
			lpGUID->Data4[1],
			lpGUID->Data4[2],
			lpGUID->Data4[3],
			lpGUID->Data4[4],
			lpGUID->Data4[5],
			lpGUID->Data4[6],
			lpGUID->Data4[7]);
	}

	std::wstring GUIDToStringAndName(_In_ GUID guid) { return GUIDToStringAndName(&guid); }

	std::wstring GUIDToStringAndName(_In_opt_ LPCGUID lpGUID)
	{
		auto szGUID = GUIDToString(lpGUID);

		szGUID += L" = "; // STRING_OK

		if (lpGUID)
		{
			for (const auto& guid : PropGuidArray)
			{
				if (IsEqualGUID(*lpGUID, *guid.lpGuid))
				{
					return szGUID + guid.lpszName;
				}
			}
		}

		return szGUID + strings::loadstring(IDS_UNKNOWNGUID);
	}

	LPCGUID GUIDNameToGUID(_In_ const std::wstring& szGUID, bool bByteSwapped)
	{
		LPGUID lpGuidRet = nullptr;
		LPCGUID lpGUID = nullptr;
		GUID guid = {};

		// Try the GUID like PS_* first
		for (const auto& propGuid : PropGuidArray)
		{
			if (0 == lstrcmpiW(szGUID.c_str(), propGuid.lpszName))
			{
				lpGUID = propGuid.lpGuid;
				break;
			}
		}

		if (!lpGUID) // no match - try it like a guid {}
		{
			guid = StringToGUID(szGUID, bByteSwapped);
			if (guid != GUID_NULL)
			{
				lpGUID = &guid;
			}
		}

		if (lpGUID)
		{
			lpGuidRet = new GUID;
			if (lpGuidRet)
			{
				memcpy(lpGuidRet, lpGUID, sizeof(GUID));
			}
		}

		return lpGuidRet;
	}

	_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID) { return StringToGUID(szGUID, false); }

	_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID, bool bByteSwapped)
	{
		auto guid = GUID_NULL;
		if (szGUID.empty()) return guid;

		auto bin = strings::HexStringToBin(szGUID, sizeof(GUID));
		if (bin.size() == sizeof(GUID))
		{
			if (bByteSwapped)
			{
				memcpy(&guid, bin.data(), sizeof(GUID));
			}
			else
			{
				guid.Data1 = (bin[0] << 24) + (bin[1] << 16) + (bin[2] << 8) + (bin[3] << 0);
				guid.Data2 = (bin[4] << 8) + (bin[5] << 0);
				guid.Data3 = (bin[6] << 8) + (bin[7] << 0);
				memcpy(&guid.Data4, bin.data() + 8, 8);
			}
		}

		return guid;
	}
} // namespace guid
