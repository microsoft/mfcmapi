#include <StdAfx.h>
#include <Interpret/SmartView/SIDBin.h>
#include <UI/MySecInfo.h>

namespace smartview
{
	void SIDBin::Parse()
	{
		const PSID SidStart = const_cast<LPBYTE>(m_Parser.GetCurrentAddress());
		const auto cbSid = m_Parser.RemainingBytes();
		m_Parser.Advance(cbSid);

		if (SidStart &&
			cbSid >= sizeof(SID) - sizeof(DWORD) + sizeof(DWORD) * static_cast<PISID>(SidStart)->SubAuthorityCount &&
			IsValidSid(SidStart))
		{
			auto sidAccount = sid::LookupAccountSid(SidStart);
			m_lpSidDomain = sidAccount.domain;
			m_lpSidName = sidAccount.name;
			m_lpStringSid = sid::GetTextualSid(SidStart);
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