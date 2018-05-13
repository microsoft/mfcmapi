#include "stdafx.h"
#include <Interpret/Sid.h>
#include <Interpret/InterpretProp2.h>
#include <Interpret/ExtraPropTags.h>

_Check_return_ wstring GetTextualSid(_In_ PSID pSid)
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
	auto TextualSid = format(L"S-%lu-", SID_REVISION); // STRING_OK

													   // Add SID identifier authority to the string.
	if (psia->Value[0] != 0 || psia->Value[1] != 0)
	{
		TextualSid += format(
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
		TextualSid += format(
			L"%lu", // STRING_OK
			static_cast<ULONG>(psia->Value[4] << 8) +
			static_cast<ULONG>(psia->Value[5]) +
			static_cast<ULONG>(psia->Value[3] << 16) +
			static_cast<ULONG>(psia->Value[2] << 24));
	}

	// Add SID subauthorities to the string.
	if (lpSubAuthoritiesCount)
	{
		for (DWORD dwCounter = 0; dwCounter < *lpSubAuthoritiesCount; dwCounter++)
		{
			TextualSid += format(
				L"-%lu", // STRING_OK
				*GetSidSubAuthority(pSid, dwCounter));
		}
	}

	return TextualSid;
}

wstring ACEToString(_In_ void* pACE, eAceType acetype)
{
	auto hRes = S_OK;
	vector<wstring> aceString;
	ACCESS_MASK Mask = 0;
	DWORD Flags = 0;
	GUID ObjectType = { 0 };
	GUID InheritedObjectType = { 0 };
	SID* SidStart = nullptr;
	auto bObjectFound = false;

	if (!pACE) return L"";

	const auto AceType = static_cast<PACE_HEADER>(pACE)->AceType;
	const auto AceFlags = static_cast<PACE_HEADER>(pACE)->AceFlags;

	/* Check type of ACE */
	switch (AceType)
	{
	case ACCESS_ALLOWED_ACE_TYPE:
		Mask = static_cast<ACCESS_ALLOWED_ACE *>(pACE)->Mask;
		SidStart = reinterpret_cast<SID *>(&static_cast<ACCESS_ALLOWED_ACE *>(pACE)->SidStart);
		break;
	case ACCESS_DENIED_ACE_TYPE:
		Mask = static_cast<ACCESS_DENIED_ACE *>(pACE)->Mask;
		SidStart = reinterpret_cast<SID *>(&static_cast<ACCESS_DENIED_ACE *>(pACE)->SidStart);
		break;
	case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
		Mask = static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->Mask;
		Flags = static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->Flags;
		ObjectType = static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->ObjectType;
		InheritedObjectType = static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->InheritedObjectType;
		SidStart = reinterpret_cast<SID *>(&static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->SidStart);
		bObjectFound = true;
		break;
	case ACCESS_DENIED_OBJECT_ACE_TYPE:
		Mask = static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->Mask;
		Flags = static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->Flags;
		ObjectType = static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->ObjectType;
		InheritedObjectType = static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->InheritedObjectType;
		SidStart = reinterpret_cast<SID *>(&static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->SidStart);
		bObjectFound = true;
		break;
	}

	DWORD dwSidName = 0;
	DWORD dwSidDomain = 0;
	SID_NAME_USE SidNameUse;

	WC_B(LookupAccountSidW(
		nullptr,
		SidStart,
		nullptr,
		&dwSidName,
		nullptr,
		&dwSidDomain,
		&SidNameUse));
	hRes = S_OK;

	LPWSTR lpSidName = nullptr;
	LPWSTR lpSidDomain = nullptr;

#pragma warning(push)
#pragma warning(disable:6211)
	if (dwSidName) lpSidName = new WCHAR[dwSidName];
	if (dwSidDomain) lpSidDomain = new WCHAR[dwSidDomain];
#pragma warning(pop)

	// Only make the call if we got something to get
	if (lpSidName || lpSidDomain)
	{
		WC_B(LookupAccountSidW(
			nullptr,
			SidStart,
			lpSidName,
			&dwSidName,
			lpSidDomain,
			&dwSidDomain,
			&SidNameUse));
	}

	auto lpStringSid = GetTextualSid(SidStart);
	auto szAceType = InterpretFlags(flagACEType, AceType);
	auto szAceFlags = InterpretFlags(flagACEFlag, AceFlags);
	wstring szAceMask;

	switch (acetype)
	{
	case acetypeContainer:
		szAceMask = InterpretFlags(flagACEMaskContainer, Mask);
		break;
	case acetypeMessage:
		szAceMask = InterpretFlags(flagACEMaskNonContainer, Mask);
		break;
	case acetypeFreeBusy:
		szAceMask = InterpretFlags(flagACEMaskFreeBusy, Mask);
		break;
	};

	auto szDomain = lpSidDomain ? lpSidDomain : formatmessage(IDS_NODOMAIN);
	auto szName = lpSidName ? lpSidName : formatmessage(IDS_NONAME);
	auto szSID = GetTextualSid(SidStart);
	if (szSID.empty()) szSID = formatmessage(IDS_NOSID);
	delete[] lpSidDomain;
	delete[] lpSidName;

	aceString.push_back(formatmessage(
		IDS_SIDACCOUNT,
		szDomain.c_str(),
		szName.c_str(),
		szSID.c_str(),
		AceType, szAceType.c_str(),
		AceFlags, szAceFlags.c_str(),
		Mask, szAceMask.c_str()));

	if (bObjectFound)
	{
		aceString.push_back(formatmessage(IDS_SIDOBJECTYPE));
		aceString.push_back(GUIDToStringAndName(&ObjectType));
		aceString.push_back(formatmessage(IDS_SIDINHERITEDOBJECTYPE));
		aceString.push_back(GUIDToStringAndName(&InheritedObjectType));
		aceString.push_back(formatmessage(IDS_SIDFLAGS, Flags));
	}

	return join(aceString, L"\r\n");
}

_Check_return_ HRESULT SDToString(_In_count_(cbBuf) const BYTE* lpBuf, size_t cbBuf, eAceType acetype, _In_ wstring& SDString, _In_ wstring& sdInfo)
{
	auto hRes = S_OK;
	BOOL bValidDACL = false;
	PACL pACL = nullptr;
	BOOL bDACLDefaulted = false;

	if (!lpBuf) return MAPI_E_NOT_FOUND;

	const auto pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(lpBuf);

	if (CbSecurityDescriptorHeader(lpBuf) > cbBuf || !IsValidSecurityDescriptor(pSecurityDescriptor))
	{
		SDString = formatmessage(IDS_INVALIDSD);
		return S_OK;
	}

	sdInfo = InterpretFlags(flagSecurityInfo, SECURITY_INFORMATION_OF(lpBuf));

	EC_B(GetSecurityDescriptorDacl(
		pSecurityDescriptor,
		&bValidDACL,
		&pACL,
		&bDACLDefaulted));
	if (bValidDACL && pACL)
	{
		ACL_SIZE_INFORMATION ACLSizeInfo = { 0 };
		EC_B(GetAclInformation(
			pACL,
			&ACLSizeInfo,
			sizeof ACLSizeInfo,
			AclSizeInformation));

		vector<wstring> sdString;
		for (DWORD i = 0; i < ACLSizeInfo.AceCount; i++)
		{
			void* pACE = nullptr;

			EC_B(GetAce(pACL, i, &pACE));

			if (pACE)
			{
				sdString.push_back(ACEToString(pACE, acetype));
			}
		}

		SDString = join(sdString, L"\r\n");
	}

	return hRes;
}