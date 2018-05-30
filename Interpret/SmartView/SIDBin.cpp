#include <StdAfx.h>
#include <Interpret/SmartView/SIDBin.h>
#include <UI/MySecInfo.h>

namespace smartview
{
	void SIDBin::Parse()
	{
		auto hRes = S_OK;
		const PSID SidStart = const_cast<LPBYTE>(m_Parser.GetCurrentAddress());
		const auto cbSid = m_Parser.RemainingBytes();
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
				const auto dwErr = GetLastError();
				hRes = HRESULT_FROM_WIN32(dwErr);
				if (ERROR_NONE_MAPPED != dwErr &&
					STRSAFE_E_INSUFFICIENT_BUFFER != hRes)
				{
					error::LogFunctionCall(hRes, NULL, false, false, true, dwErr, "LookupAccountSid", __FILE__, __LINE__);
				}

				hRes = S_OK;
			}

			const auto lpSidName = dwSidName ? new WCHAR[dwSidName] : nullptr;
			const auto lpSidDomain = dwSidDomain ? new WCHAR[dwSidDomain] : nullptr;

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
				if (lpSidName) m_lpSidName = lpSidName;
				if (lpSidDomain) m_lpSidDomain = lpSidDomain;
				delete[] lpSidName;
				delete[] lpSidDomain;
			}

			m_lpStringSid = GetTextualSid(SidStart);
		}
	}

	_Check_return_ std::wstring SIDBin::ToStringInternal()
	{
		auto szDomain = !m_lpSidDomain.empty() ? m_lpSidDomain : strings::formatmessage(IDS_NODOMAIN);
		auto szName = !m_lpSidName.empty() ? m_lpSidName : strings::formatmessage(IDS_NONAME);
		auto szSID = !m_lpStringSid.empty() ? m_lpStringSid : strings::formatmessage(IDS_NOSID);

		return strings::formatmessage(IDS_SIDHEADER, szDomain.c_str(), szName.c_str(), szSID.c_str());
	}
}