#include <core/stdafx.h>
#include <core/smartview/SIDBin.h>
#include <core/utility/strings.h>
#include <core/interpret/sid.h>

namespace smartview
{
	void SIDBin::Parse() { m_SIDbin = blockBytes::parse(m_Parser, m_Parser->getSize()); }

	void SIDBin::ParseBlocks()
	{
		if (m_SIDbin)
		{
			auto sidAccount = sid::LookupAccountSid(*m_SIDbin);
			auto sidString = sid::GetTextualSid(*m_SIDbin);

			setRoot(L"SID: \r\n");
			addHeader(L"User: %1!ws!\\%2!ws!\r\n", sidAccount.getDomain().c_str(), sidAccount.getName().c_str());

			if (sidString.empty()) sidString = strings::formatmessage(IDS_NOSID);
			addChild(m_SIDbin, L"Textual SID: %1!ws!", sidString.c_str());
		}
	}
} // namespace smartview