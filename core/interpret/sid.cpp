#include <core/stdafx.h>
#include <core/interpret/sid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <core/utility/error.h>
#include <core/smartview/SmartView.h>
#include <core/addin/mfcmapi.h>
#include <core/smartview/SD/NTSD.h>

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

	_Check_return_ SidAccount LookupAccountSid(std::vector<BYTE> buf)
	{
		const auto subAuthorityCount = buf.size() >= 2 ? buf[1] : 0;
		if (buf.size() < sizeof(SID) - sizeof(DWORD) + sizeof(DWORD) * subAuthorityCount) return {};

		PSID SidStart = buf.data();
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

	_Check_return_ std::wstring NTSDToString(const std::vector<BYTE>& buf, aceType acetype)
	{
		const std::shared_ptr<smartview::block> svp = std::make_shared<smartview::NTSD>(acetype);
		if (svp)
		{
			svp->parse(std::make_shared<smartview::binaryParser>(buf), true);
			return svp->toString();
		}

		return {};
	}
} // namespace sid