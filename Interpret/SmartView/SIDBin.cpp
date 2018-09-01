#include <StdAfx.h>
#include <Interpret/SmartView/SIDBin.h>
#include <UI/MySecInfo.h>

namespace smartview
{
	void SIDBin::Parse() { m_SIDbin = m_Parser.GetBlockBYTES(m_Parser.RemainingBytes()); }

	void SIDBin::ParseBlocks()
	{
		auto piSid = reinterpret_cast<PISID>(const_cast<LPBYTE>(m_SIDbin.data()));
		auto sidAccount = sid::SidAccount{};
		auto sidString = std::wstring{};
		if (!m_SIDbin.empty() &&
			m_SIDbin.size() >= sizeof(SID) - sizeof(DWORD) + sizeof(DWORD) * piSid->SubAuthorityCount &&
			IsValidSid(piSid))
		{
			sidAccount = sid::LookupAccountSid(piSid);
			sidString = sid::GetTextualSid(piSid);
		}

		addHeader(L"SID: \r\n");
		addBlock(m_SIDbin, L"User: %1!ws!\\%2!ws!\r\n", sidAccount.getDomain().c_str(), sidAccount.getName().c_str());

		if (sidString.empty()) sidString = strings::formatmessage(IDS_NOSID);
		addBlock(m_SIDbin, L"Textual SID: %1!ws!", sidString.c_str());
	}
}