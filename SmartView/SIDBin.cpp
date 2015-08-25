#include "stdafx.h"
#include "..\stdafx.h"
#include "SIDBin.h"
#include "..\String.h"
#include "..\MySecInfo.h"

SIDBin::SIDBin(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_lpSidName = NULL;
	m_lpSidDomain = NULL;
}

SIDBin::~SIDBin()
{
	delete[] m_lpSidName;
	delete[] m_lpSidDomain;
}

void SIDBin::Parse()
{
	HRESULT hRes = S_OK;
	PSID SidStart = m_Parser.GetCurrentAddress();
	size_t cbSid = m_Parser.RemainingBytes();
	m_Parser.Advance(cbSid);

	if (SidStart &&
		cbSid >= sizeof(SID) - sizeof(DWORD) + sizeof(DWORD)* ((PISID)SidStart)->SubAuthorityCount &&
		IsValidSid(SidStart))
	{
		DWORD dwSidName = 0;
		DWORD dwSidDomain = 0;
		SID_NAME_USE SidNameUse;

		if (!LookupAccountSid(
			NULL,
			SidStart,
			NULL,
			&dwSidName,
			NULL,
			&dwSidDomain,
			&SidNameUse))
		{
			DWORD dwErr = GetLastError();
			hRes = HRESULT_FROM_WIN32(dwErr);
			if (ERROR_NONE_MAPPED != dwErr &&
				STRSAFE_E_INSUFFICIENT_BUFFER != hRes)
			{
				LogFunctionCall(hRes, NULL, false, false, true, dwErr, "LookupAccountSid", __FILE__, __LINE__);
			}

			hRes = S_OK;
		}

#pragma warning(push)
#pragma warning(disable:6211)
		if (dwSidName) m_lpSidName = new TCHAR[dwSidName];
		if (dwSidDomain) m_lpSidDomain = new TCHAR[dwSidDomain];
#pragma warning(pop)

		// Only make the call if we got something to get
		if (m_lpSidName || m_lpSidDomain)
		{
			WC_B(LookupAccountSid(
				NULL,
				SidStart,
				m_lpSidName,
				&dwSidName,
				m_lpSidDomain,
				&dwSidDomain,
				&SidNameUse));
			hRes = S_OK;
		}

		m_lpStringSid = GetTextualSid(SidStart);
	}
}

_Check_return_ wstring SIDBin::ToStringInternal()
{
	wstring szDomain = m_lpSidDomain ? LPCTSTRToWstring(m_lpSidDomain) : formatmessage(IDS_NODOMAIN);
	wstring szName = m_lpSidName ? LPCTSTRToWstring(m_lpSidName) : formatmessage(IDS_NONAME);
	wstring szSID = !m_lpStringSid.empty() ? m_lpStringSid : formatmessage(IDS_NOSID);

	return formatmessage(IDS_SIDHEADER, szDomain.c_str(), szName.c_str(), szSID.c_str());
}