#include <core/stdafx.h>
#include <core/smartview/SD/SIDBin.h>
#include <core/utility/strings.h>
#include <core/interpret/sid.h>

namespace smartview
{
	void SIDBin::parse() { m_SIDbin = blockBytes::parse(parser, parser->getSize()); }

	void SIDBin::parseBlocks()
	{
		if (m_SIDbin)
		{
			auto sidAccount = sid::LookupAccountSid(*m_SIDbin);
			auto sidString = sid::GetTextualSid(*m_SIDbin);

			setText(L"SID");
			addHeader(L"User: %1!ws!\\%2!ws!", sidAccount.getDomain().c_str(), sidAccount.getName().c_str());

			if (sidString.empty()) sidString = strings::formatmessage(IDS_NOSID);
			addChild(m_SIDbin, L"Textual SID: %1!ws!", sidString.c_str());
		}
	}
} // namespace smartview