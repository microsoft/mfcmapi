#include "stdafx.h"
#include "SIDBin.h"
#include "MySecInfo.h"

void SIDBin::Parse()
{
	auto hRes = S_OK;
	PSID SidStart = m_Parser.GetCurrentAddress();
	auto cbSid = m_Parser.RemainingBytes();
	m_Parser.Advance(cbSid);

	if (SidStart &&
		cbSid >= sizeof(SID) - sizeof(DWORD) + sizeof(DWORD)* static_cast<PISID>(SidStart)->SubAuthorityCount &&
		IsValidSid(SidStart))
	{
		DWORD dwSidName = 0;
		DWORD dwSidDomain = 0;
		SID_NAME_USE SidNameUse;

		if (!LookupAccountSidW(
			nullptr,
			SidStart,
			nullptr,
			&dwSidName,
			nullptr,
			&dwSidDomain,
			&SidNameUse))
		{
			auto dwErr = GetLastError();
			hRes = HRESULT_FROM_WIN32(dwErr);
			if (ERROR_NONE_MAPPED != dwErr &&
				STRSAFE_E_INSUFFICIENT_BUFFER != hRes)
			{
				LogFunctionCall(hRes, NULL, false, false, true, dwErr, "LookupAccountSid", __FILE__, __LINE__);
			}

			hRes = S_OK;
		}

		auto lpSidName = dwSidName ? new WCHAR[dwSidName] : nullptr;
		auto lpSidDomain = dwSidDomain ? new WCHAR[dwSidDomain] : nullptr;

		// Only make the call if we got something to get
		if (lpSidName || lpSidDomain)
		{
			WC_B(LookupAccountSidW(
				NULL,
				SidStart,
				lpSidName,
				&dwSidName,
				lpSidDomain,
				&dwSidDomain,
				&SidNameUse));
			if (lpSidName) m_lpSidName = lpSidName;
			if (lpSidDomain) m_lpSidDomain = lpSidDomain;
			delete[] lpSidName;
			delete[] lpSidDomain;
		}

		m_lpStringSid = GetTextualSid(SidStart);
	}
}

_Check_return_ wstring SIDBin::ToStringInternal()
{
	auto szDomain = !m_lpSidDomain.empty() ? m_lpSidDomain : formatmessage(IDS_NODOMAIN);
	auto szName = !m_lpSidName.empty() ? m_lpSidName : formatmessage(IDS_NONAME);
	auto szSID = !m_lpStringSid.empty() ? m_lpStringSid : formatmessage(IDS_NOSID);

	return formatmessage(IDS_SIDHEADER, szDomain.c_str(), szName.c_str(), szSID.c_str());
}