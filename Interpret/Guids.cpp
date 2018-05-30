#include <StdAfx.h>

#pragma region "USES_IID definitions"
#define INITGUID
#define USES_IID_IMsgStore
#define USES_IID_IMAPIProgress
#define USES_IID_IMAPIAdviseSink
#define USES_IID_IMAPIMessageSite
#define USES_IID_IMAPIViewContext
#define USES_IID_IMAPIViewAdviseSink
#define USES_IID_IMAPITable
#define USES_IID_IMAPIFolder
#define USES_IID_IMAPIContainer
#define USES_IID_IMAPIForm
#define USES_IID_IMAPIProp
#define USES_IID_IPersistMessage
#define USES_IID_IMessage
#define USES_IID_IAttachment
#define USES_IID_IUnknown
#define USES_IID_IABContainer
#define USES_IID_IDistList
#define USES_IID_IStreamDocfile
#define USES_PS_MAPI
#define USES_PS_PUBLIC_STRINGS
#define USES_PS_ROUTING_EMAIL_ADDRESSES
#define USES_PS_ROUTING_ADDRTYPE
#define USES_PS_ROUTING_DISPLAY_NAME
#define USES_PS_ROUTING_ENTRYID
#define USES_PS_ROUTING_SEARCH_KEY
#define USES_MUID_PROFILE_INSTANCE
#define USES_IID_IMAPIClientShutdown
#define USES_IID_IMAPIProviderShutdown
#define USES_IID_IMsgServiceAdmin2
#define USES_IID_IMessageRaw
#pragma endregion

#include <initguid.h>
#include <MAPIGuid.h>
#include <MAPIAux.h>

#ifdef EDKGUID_INCLUDED
#undef EDKGUID_INCLUDED
#endif
#include <EdkGuid.h>

#include <Interpret/Guids.h>

namespace guid
{
	std::wstring GUIDToString(_In_opt_ LPCGUID lpGUID)
	{
		GUID nullGUID = { 0 };

		if (!lpGUID)
		{
			lpGUID = &nullGUID;
		}

		return strings::format(L"{%.8X-%.4X-%.4X-%.2X%.2X-%.2X%.2X%.2X%.2X%.2X%.2X}", // STRING_OK
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
		GUID guid = { 0 };

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

	_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID)
	{
		return StringToGUID(szGUID, false);
	}

	_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID, bool bByteSwapped)
	{
		auto guid = GUID_NULL;
		if (szGUID.empty()) return guid;

		auto bin = strings::HexStringToBin(szGUID, sizeof(GUID));
		if (bin.size() == sizeof(GUID))
		{
			memcpy(&guid, bin.data(), sizeof(GUID));

			// Note that we get the bByteSwapped behavior by default. We have to work to get the 'normal' behavior
			if (!bByteSwapped)
			{
				const auto lpByte = reinterpret_cast<LPBYTE>(&guid);
				auto bByte = lpByte[0];
				lpByte[0] = lpByte[3];
				lpByte[3] = bByte;
				bByte = lpByte[1];
				lpByte[1] = lpByte[2];
				lpByte[2] = bByte;
			}
		}

		return guid;
	}
}
