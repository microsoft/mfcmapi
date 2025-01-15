#include <core/stdafx.h>
#include <core/interpret/sid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <core/utility/error.h>
#include <core/smartview/SmartView.h>
#include <core/addin/mfcmapi.h>

namespace sid
{
	_Check_return_ std::wstring SidAccount::getDomain() const
	{
		return !domain.empty() ? domain : strings::formatmessage(IDS_NODOMAIN);
	}

	_Check_return_ std::wstring SidAccount::getName() const
	{
		return !name.empty() ? name : strings::formatmessage(IDS_NONAME);
	}

	_Check_return_ std::wstring LookupIdentifierAuthority(const SID_IDENTIFIER_AUTHORITY& authority)
	{
		static const auto authorityLookupTable = std::vector<std::pair<SID_IDENTIFIER_AUTHORITY, std::wstring>>{
			{SECURITY_NULL_SID_AUTHORITY, L"SECURITY_NULL_SID_AUTHORITY"},
			{SECURITY_WORLD_SID_AUTHORITY, L"SECURITY_WORLD_SID_AUTHORITY"},
			{SECURITY_LOCAL_SID_AUTHORITY, L"SECURITY_LOCAL_SID_AUTHORITY"},
			{SECURITY_CREATOR_SID_AUTHORITY, L"SECURITY_CREATOR_SID_AUTHORITY"},
			{SECURITY_NON_UNIQUE_AUTHORITY, L"SECURITY_NON_UNIQUE_AUTHORITY"},
			{SECURITY_RESOURCE_MANAGER_AUTHORITY, L"SECURITY_RESOURCE_MANAGER_AUTHORITY"},
			{SECURITY_NT_AUTHORITY, L"SECURITY_NT_AUTHORITY"},
			{SECURITY_APP_PACKAGE_AUTHORITY, L"SECURITY_APP_PACKAGE_AUTHORITY"},
			{SECURITY_MANDATORY_LABEL_AUTHORITY, L"SECURITY_MANDATORY_LABEL_AUTHORITY"},
			{SECURITY_SCOPED_POLICY_ID_AUTHORITY, L"SECURITY_SCOPED_POLICY_ID_AUTHORITY"},
			{SECURITY_AUTHENTICATION_AUTHORITY, L"SECURITY_AUTHENTICATION_AUTHORITY"},
			{SECURITY_PROCESS_TRUST_AUTHORITY, L"SECURITY_PROCESS_TRUST_AUTHORITY"},
		};

		for (const auto& entry : authorityLookupTable)
		{
			if (std::memcmp(&authority, &entry.first, sizeof(SID_IDENTIFIER_AUTHORITY)) == 0)
			{
				return entry.second;
			}
		}

		return IdentifierAuthorityToString(authority);
	}

	_Check_return_ std::wstring IdentifierAuthorityToString(const SID_IDENTIFIER_AUTHORITY& authority)
	{
		if (authority.Value[0] != 0 || authority.Value[1] != 0)
		{
			return strings::format(
				L"%02hx%02hx%02hx%02hx%02hx%02hx", // STRING_OK
				static_cast<USHORT>(authority.Value[0]),
				static_cast<USHORT>(authority.Value[1]),
				static_cast<USHORT>(authority.Value[2]),
				static_cast<USHORT>(authority.Value[3]),
				static_cast<USHORT>(authority.Value[4]),
				static_cast<USHORT>(authority.Value[5]));
		}
		else
		{
			return strings::format(
				L"%lu", // STRING_OK
				static_cast<ULONG>(authority.Value[4] << 8) + static_cast<ULONG>(authority.Value[5]) +
					static_cast<ULONG>(authority.Value[3] << 16) + static_cast<ULONG>(authority.Value[2] << 24));
		}
	}

	// [MS-DTYP] 2.4.2.2 SID--Packet Representation
	// https://msdn.microsoft.com/en-us/library/gg465313.aspx
	_Check_return_ std::wstring GetTextualSid(_In_opt_ PSID pSid)
	{
		// Validate the binary SID.
		if (!pSid || !IsValidSid(pSid)) return {};

		// Get the identifier authority value from the SID.
		const auto psia = GetSidIdentifierAuthority(pSid);

		// Get the number of subauthorities in the SID.
		const auto lpSubAuthoritiesCount = GetSidSubAuthorityCount(pSid);

		// Compute the buffer length.
		// S-SID_REVISION- + IdentifierAuthority- + subauthorities- + NULL
		// Add 'S' prefix and revision number to the string.
		auto TextualSid = strings::format(L"S-%lu-", SID_REVISION); // STRING_OK

		// Add SID identifier authority to the string.
		TextualSid += IdentifierAuthorityToString(*psia);

		// Add SID subauthorities to the string.
		if (lpSubAuthoritiesCount)
		{
			for (DWORD dwCounter = 0; dwCounter < *lpSubAuthoritiesCount; dwCounter++)
			{
				if (pSid)
				{
					TextualSid += strings::format(
						L"-%lu", // STRING_OK
						*GetSidSubAuthority(pSid, dwCounter));
				}
			}
		}

		return TextualSid;
	}

	_Check_return_ std::wstring GetTextualSid(std::vector<BYTE> buf)
	{
		const auto subAuthorityCount = buf.size() >= 2 ? buf[1] : 0;
		if (buf.size() < sizeof(SID) - sizeof(DWORD) + sizeof(DWORD) * subAuthorityCount) return {};

		return GetTextualSid(buf.data());
	}

	_Check_return_ SidAccount LookupAccountSid(PSID SidStart)
	{
		if (!IsValidSid(SidStart)) return {};

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
					HRESULT_FROM_WIN32(dwErr), {}, false, false, true, dwErr, "LookupAccountSid", __FILE__, __LINE__);
			}
		}

		auto sidNameBuf = std::vector<wchar_t>(cchSidName);
		auto sidDomainBuf = std::vector<wchar_t>(cchSidDomain);
		WC_B_S(LookupAccountSidW(
			nullptr,
			SidStart,
			cchSidName ? &sidNameBuf[0] : nullptr,
			&cchSidName,
			cchSidDomain ? &sidDomainBuf[0] : nullptr,
			&cchSidDomain,
			&SidNameUse));

		if (cchSidName && sidNameBuf.back() == L'\0') sidNameBuf.pop_back();
		if (cchSidDomain && sidDomainBuf.back() == L'\0') sidDomainBuf.pop_back();

		return SidAccount{
			std::wstring(sidDomainBuf.begin(), sidDomainBuf.end()), std::wstring(sidNameBuf.begin(), sidNameBuf.end())};
	}

	_Check_return_ SidAccount LookupAccountSid(std::vector<BYTE> buf)
	{
		const auto subAuthorityCount = buf.size() >= 2 ? buf[1] : 0;
		if (buf.size() < sizeof(SID) - sizeof(DWORD) + sizeof(DWORD) * subAuthorityCount) return {};

		return LookupAccountSid(buf.data());
	}

	std::wstring ACEToString(const std::vector<BYTE>& buf, aceType acetype)
	{
		parserType parser = parserType::ACEMESSAGE;
		switch (acetype)
		{
		case aceType::Container:
			parser = parserType::ACECONTAINER;
			break;
		case aceType::Message:
			parser = parserType::ACEMESSAGE;
			break;
		case aceType::FreeBusy:
			parser = parserType::ACEFB;
			break;
		}

		return smartview::InterpretBinary(
				   {static_cast<ULONG>(buf.size()), const_cast<LPBYTE>(buf.data())}, parser, nullptr)
			->toString();
	}

	std::wstring ACEToString(_In_opt_ void* pACE, aceType acetype)
	{
		std::vector<std::wstring> aceString;
		PACE_HEADER pAceHeader = static_cast<PACE_HEADER>(pACE);
		ACCESS_MASK Mask = 0;
		DWORD Flags = 0;
		GUID ObjectType = {};
		GUID InheritedObjectType = {};
		SID* SidStart = nullptr;
		auto bObjectFound = false;

		if (!pACE) return L"";

		/* Check type of ACE */
		switch (pAceHeader->AceType)
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
		auto szAceType = flags::InterpretFlags(flagACEType, pAceHeader->AceType);
		auto szAceFlags = flags::InterpretFlags(flagACEFlag, pAceHeader->AceFlags);
		auto szAceMask = std::wstring{};

		switch (acetype)
		{
		case aceType::Container:
			szAceMask = flags::InterpretFlags(flagACEMaskContainer, Mask);
			break;
		case aceType::Message:
			szAceMask = flags::InterpretFlags(flagACEMaskNonContainer, Mask);
			break;
		case aceType::FreeBusy:
			szAceMask = flags::InterpretFlags(flagACEMaskFreeBusy, Mask);
			break;
		};

		auto sidAccount = sid::LookupAccountSid(SidStart);

		auto szSID = GetTextualSid(SidStart);
		if (szSID.empty()) szSID = strings::formatmessage(IDS_NOSID);

		aceString.push_back(strings::formatmessage(
			IDS_SIDACCOUNT,
			sidAccount.getDomain().c_str(),
			sidAccount.getName().c_str(),
			pAceHeader->AceType,
			szAceType.c_str(),
			pAceHeader->AceFlags,
			szAceFlags.c_str(),
			Mask,
			szAceMask.c_str(),
			pAceHeader->AceSize,
			szSID.c_str()));

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

	_Check_return_ bool IsValidSecurityDescriptorEx(const std::vector<BYTE>& buf) noexcept
	{
		try
		{
			if (buf.empty() || buf.size() < 2 * sizeof(DWORD)) return false;
			if (CbSecurityDescriptorHeader(buf.data()) >= buf.size()) return false;
			const auto pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(buf.data());
			return IsValidSecurityDescriptor(pSecurityDescriptor);
		} catch (...)
		{
			return false;
		}
	}

	_Check_return_ SecurityDescriptor NTSDToString(const std::vector<BYTE>& buf, aceType acetype)
	{
		if (!IsValidSecurityDescriptorEx(buf))
			return SecurityDescriptor{strings::formatmessage(IDS_INVALIDSD), strings::emptystring};
		const auto pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(buf.data());
		const auto cbSecurityDescriptor = buf.size() - CbSecurityDescriptorHeader(buf.data());
		const auto sdVector = std::vector<BYTE>(pSecurityDescriptor, pSecurityDescriptor + cbSecurityDescriptor);
		const auto sdString = SDToString(sdVector, acetype);

		return SecurityDescriptor{
			sdString, flags::InterpretFlags(flagSecurityInfo, SECURITY_INFORMATION_OF(buf.data()))};
	}

	_Check_return_ std::wstring SDToString(const std::vector<BYTE>& buf, aceType acetype)
	{
		const auto pSecurityDescriptor = const_cast<LPBYTE>(buf.data());
		if (!IsValidSecurityDescriptor(pSecurityDescriptor)) return {};

		auto bValidDACL = static_cast<BOOL>(false);
		auto pACL = PACL{};
		auto bDACLDefaulted = static_cast<BOOL>(false);
		auto sdString = std::vector<std::wstring>{};
		EC_B_S(GetSecurityDescriptorDacl(pSecurityDescriptor, &bValidDACL, &pACL, &bDACLDefaulted));
		if (bValidDACL && pACL)
		{
			auto ACLSizeInfo = ACL_SIZE_INFORMATION{};
			EC_B_S(GetAclInformation(pACL, &ACLSizeInfo, sizeof ACLSizeInfo, AclSizeInformation));

			for (DWORD i = 0; i < ACLSizeInfo.AceCount; i++)
			{
				auto pACE = LPVOID{};

				WC_B_S(GetAce(pACL, i, &pACE));
				if (pACE)
				{
					// TODO: Replace this with the counted buffer variant
					sdString.push_back(ACEToString(pACE, acetype));
				}
			}
		}

		return strings::join(sdString, L"\r\n");
	}
} // namespace sid