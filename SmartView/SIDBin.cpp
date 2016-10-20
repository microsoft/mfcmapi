#include "stdafx.h"
#include "SIDBin.h"
#include "String.h"
#include "MySecInfo.h"

SIDBin::SIDBin()
{
	m_lpSidName = nullptr;
	m_lpSidDomain = nullptr;
}

SIDBin::~SIDBin()
{
	delete[] m_lpSidName;
	delete[] m_lpSidDomain;
}

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

		if (dwSidName) m_lpSidName = new WCHAR[dwSidName];
		if (dwSidDomain) m_lpSidDomain = new WCHAR[dwSidDomain];

		// Only make the call if we got something to get
		if (m_lpSidName || m_lpSidDomain)
		{
			WC_B(LookupAccountSidW(
				NULL,
				SidStart,
				m_lpSidName,
				&dwSidName,
				m_lpSidDomain,
				&dwSidDomain,
				&SidNameUse));
		}

		m_lpStringSid = GetTextualSid(SidStart);
	}
}

_Check_return_ wstring SIDBin::ToStringInternal()
{
	auto szDomain = m_lpSidDomain ? m_lpSidDomain : formatmessage(IDS_NODOMAIN);
	auto szName = m_lpSidName ? m_lpSidName : formatmessage(IDS_NONAME);
	auto szSID = !m_lpStringSid.empty() ? m_lpStringSid : formatmessage(IDS_NOSID);

	return formatmessage(IDS_SIDHEADER, szDomain.c_str(), szName.c_str(), szSID.c_str());
}