#include <StdAfx.h>
#include <Interpret/Sid.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/GUIDArray.h>

namespace sid
{
	_Check_return_ std::wstring GetTextualSid(_In_ PSID pSid)
	{
		// Validate the binary SID.
		if (!IsValidSid(pSid)) return L"";

		// Get the identifier authority value from the SID.
		const auto psia = GetSidIdentifierAuthority(pSid);

		// Get the number of subauthorities in the SID.
		const auto lpSubAuthoritiesCount = GetSidSubAuthorityCount(pSid);

		// Compute the buffer length.
		// S-SID_REVISION- + IdentifierAuthority- + subauthorities- + NULL
		// Add 'S' prefix and revision number to the string.
		auto TextualSid = strings::format(L"S-%lu-", SID_REVISION); // STRING_OK

		// Add SID identifier authority to the string.
		if (psia->Value[0] != 0 || psia->Value[1] != 0)
		{
			TextualSid += strings::format(
				L"0x%02hx%02hx%02hx%02hx%02hx%02hx", // STRING_OK
				static_cast<USHORT>(psia->Value[0]),
				static_cast<USHORT>(psia->Value[1]),
				static_cast<USHORT>(psia->Value[2]),
				static_cast<USHORT>(psia->Value[3]),
				static_cast<USHORT>(psia->Value[4]),
				static_cast<USHORT>(psia->Value[5]));
		}
		else
		{
			TextualSid += strings::format(
				L"%lu", // STRING_OK
				static_cast<ULONG>(psia->Value[4] << 8) + static_cast<ULONG>(psia->Value[5]) +
					static_cast<ULONG>(psia->Value[3] << 16) + static_cast<ULONG>(psia->Value[2] << 24));
		}

		// Add SID subauthorities to the string.
		if (lpSubAuthoritiesCount)
		{
			for (DWORD dwCounter = 0; dwCounter < *lpSubAuthoritiesCount; dwCounter++)
			{
				TextualSid += strings::format(
					L"-%lu", // STRING_OK
					*GetSidSubAuthority(pSid, dwCounter));
			}
		}

		return TextualSid;
	}

	_Check_return_ SidAccount LookupAccountSid(PSID SidStart)
	{
		// TODO: Make use of SidNameUse information
		auto cchSidName = DWORD{};
		auto cchSidDomain = DWORD{};
		auto SidNameUse = SID_NAME_USE{};

		if (!LookupAccountSidW(nullptr, SidStart, nullptr, &cchSidName, nullptr, &cchSidDomain, &SidNameUse))
		{
			const auto dwErr = GetLastError();
			if (dwErr != ERROR_NONE_MAPPED && dwErr != ERROR_INSUFFICIENT_BUFFER)
			{
				error::LogFunctionCall(
					HRESULT_FROM_WIN32(dwErr), NULL, false, false, true, dwErr, "LookupAccountSid", __FILE__, __LINE__);
			}
		}

		auto sidNameBuf = std::vector<wchar_t>();
		sidNameBuf.resize(cchSidName);
		auto sidDomainBuf = std::vector<wchar_t>();
		sidDomainBuf.resize(cchSidDomain);
		WC_B_S(LookupAccountSidW(
			nullptr,
			SidStart,
			cchSidName ? &sidNameBuf.at(0) : nullptr,
			&cchSidName,
			cchSidDomain ? &sidDomainBuf.at(0) : nullptr,
			&cchSidDomain,
			&SidNameUse));

		return SidAccount{std::wstring(sidDomainBuf.begin(), sidDomainBuf.end()),
						  std::wstring(sidNameBuf.begin(), sidNameBuf.end())};
	}

	std::wstring ACEToString(_In_ void* pACE, eAceType acetype)
	{
		std::vector<std::wstring> aceString;
		ACCESS_MASK Mask = 0;
		DWORD Flags = 0;
		GUID ObjectType = {0};
		GUID InheritedObjectType = {0};
		SID* SidStart = nullptr;
		auto bObjectFound = false;

		if (!pACE) return L"";

		const auto AceType = static_cast<PACE_HEADER>(pACE)->AceType;
		const auto AceFlags = static_cast<PACE_HEADER>(pACE)->AceFlags;

		/* Check type of ACE */
		switch (AceType)
		{
		case ACCESS_ALLOWED_ACE_TYPE:
			Mask = static_cast<ACCESS_ALLOWED_ACE*>(pACE)->Mask;
			SidStart = reinterpret_cast<SID*>(&static_cast<ACCESS_ALLOWED_ACE*>(pACE)->SidStart);
			break;
		case ACCESS_DENIED_ACE_TYPE:
			Mask = static_cast<ACCESS_DENIED_ACE*>(pACE)->Mask;
			SidStart = reinterpret_cast<SID*>(&static_cast<ACCESS_DENIED_ACE*>(pACE)->SidStart);
			break;
		case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
			Mask = static_cast<ACCESS_ALLOWED_OBJECT_ACE*>(pACE)->Mask;
			Flags = static_cast<ACCESS_ALLOWED_OBJECT_ACE*>(pACE)->Flags;
			ObjectType = static_cast<ACCESS_ALLOWED_OBJECT_ACE*>(pACE)->ObjectType;
			InheritedObjectType = static_cast<ACCESS_ALLOWED_OBJECT_ACE*>(pACE)->InheritedObjectType;
			SidStart = reinterpret_cast<SID*>(&static_cast<ACCESS_ALLOWED_OBJECT_ACE*>(pACE)->SidStart);
			bObjectFound = true;
			break;
		case ACCESS_DENIED_OBJECT_ACE_TYPE:
			Mask = static_cast<ACCESS_DENIED_OBJECT_ACE*>(pACE)->Mask;
			Flags = static_cast<ACCESS_DENIED_OBJECT_ACE*>(pACE)->Flags;
			ObjectType = static_cast<ACCESS_DENIED_OBJECT_ACE*>(pACE)->ObjectType;
			InheritedObjectType = static_cast<ACCESS_DENIED_OBJECT_ACE*>(pACE)->InheritedObjectType;
			SidStart = reinterpret_cast<SID*>(&static_cast<ACCESS_DENIED_OBJECT_ACE*>(pACE)->SidStart);
			bObjectFound = true;
			break;
		}

		auto lpStringSid = GetTextualSid(SidStart);
		auto szAceType = interpretprop::InterpretFlags(flagACEType, AceType);
		auto szAceFlags = interpretprop::InterpretFlags(flagACEFlag, AceFlags);
		std::wstring szAceMask;

		switch (acetype)
		{
		case acetypeContainer:
			szAceMask = interpretprop::InterpretFlags(flagACEMaskContainer, Mask);
			break;
		case acetypeMessage:
			szAceMask = interpretprop::InterpretFlags(flagACEMaskNonContainer, Mask);
			break;
		case acetypeFreeBusy:
			szAceMask = interpretprop::InterpretFlags(flagACEMaskFreeBusy, Mask);
			break;
		};

		auto sidAccount = sid::LookupAccountSid(SidStart);

		auto szSID = GetTextualSid(SidStart);
		if (szSID.empty()) szSID = strings::formatmessage(IDS_NOSID);

		aceString.push_back(strings::formatmessage(
			IDS_SIDACCOUNT,
			sidAccount.getDomain().c_str(),
			sidAccount.getName().c_str(),
			szSID.c_str(),
			AceType,
			szAceType.c_str(),
			AceFlags,
			szAceFlags.c_str(),
			Mask,
			szAceMask.c_str()));

		if (bObjectFound)
		{
			aceString.push_back(strings::formatmessage(IDS_SIDOBJECTYPE));
			aceString.push_back(guid::GUIDToStringAndName(&ObjectType));
			aceString.push_back(strings::formatmessage(IDS_SIDINHERITEDOBJECTYPE));
			aceString.push_back(guid::GUIDToStringAndName(&InheritedObjectType));
			aceString.push_back(strings::formatmessage(IDS_SIDFLAGS, Flags));
		}

		return strings::join(aceString, L"\r\n");
	}

	_Check_return_ SecurityDescriptor SDToString(_In_count_(cbBuf) const BYTE* lpBuf, size_t cbBuf, eAceType acetype)
	{
		if (!lpBuf) return {};

		const auto pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(lpBuf);

		if (CbSecurityDescriptorHeader(lpBuf) > cbBuf || !IsValidSecurityDescriptor(pSecurityDescriptor))
		{
			return SecurityDescriptor(strings::formatmessage(IDS_INVALIDSD), strings::emptystring);
		}

		const auto sdInfo = interpretprop::InterpretFlags(flagSecurityInfo, SECURITY_INFORMATION_OF(lpBuf));

		BOOL bValidDACL = false;
		PACL pACL = nullptr;
		BOOL bDACLDefaulted = false;
		EC_B_S(GetSecurityDescriptorDacl(pSecurityDescriptor, &bValidDACL, &pACL, &bDACLDefaulted));
		if (bValidDACL && pACL)
		{
			ACL_SIZE_INFORMATION ACLSizeInfo = {};
			EC_B_S(GetAclInformation(pACL, &ACLSizeInfo, sizeof ACLSizeInfo, AclSizeInformation));

			std::vector<std::wstring> sdString;
			for (DWORD i = 0; i < ACLSizeInfo.AceCount; i++)
			{
				void* pACE = nullptr;

				WC_B_S(GetAce(pACL, i, &pACE));
				if (pACE)
				{
					sdString.push_back(ACEToString(pACE, acetype));
				}
			}

			return SecurityDescriptor(strings::join(sdString, L"\r\n"), sdInfo);
		}

		return SecurityDescriptor(strings::emptystring, sdInfo);
	}
} // namespace sid