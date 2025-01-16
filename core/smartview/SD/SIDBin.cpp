#include <core/stdafx.h>
#include <core/smartview/SD/SIDBin.h>
#include <core/utility/strings.h>
#include <core/interpret/sid.h>

namespace smartview
{
	void SIDBin::parse()
	{
		const auto sidOffset = parser->getOffset();

		Revision = blockT<BYTE>::parse(parser);
		SubAuthorityCount = blockT<BYTE>::parse(parser);
		IdentifierAuthority = blockBytes::parse(parser, 6); // 6 bytes
		for (auto i = 0; i < SubAuthorityCount->getData(); i++)
		{
			const auto sa = blockT<DWORD>::parse(parser);
			if (!sa->isSet()) break;
			SubAuthority.push_back(sa);
		}

		const auto postSidOffset = parser->getOffset();
		parser->setOffset(sidOffset);
		m_SIDbin = blockBytes::parse(parser, postSidOffset - sidOffset);
	}

	void SIDBin::parseBlocks()
	{
		setText(L"SID");

		if (m_SIDbin)
		{
			auto sidAccount = sid::LookupAccountSid(*m_SIDbin);
			addHeader(L"User: %1!ws!\\%2!ws!", sidAccount.getDomain().c_str(), sidAccount.getName().c_str());
		}

		std::wstring TextualSid = {};
		const auto psia =
			IdentifierAuthority->isSet() ? (PSID_IDENTIFIER_AUTHORITY) (IdentifierAuthority->data()) : nullptr;

		if (psia != nullptr && SubAuthority.size() == *SubAuthorityCount)
		{
			TextualSid = strings::format(L"S-%lu-", Revision->getData());
			TextualSid += sid::IdentifierAuthorityToString(*psia);

			// Add SID subauthorities to the string.
			if (SubAuthority.size() > 0)
			{
				for (const auto& sa : SubAuthority)
				{
					TextualSid += strings::format(L"-%lu", sa->getData());
				}
			}
		}
		else
		{
			TextualSid = strings::formatmessage(IDS_NOSID);
		}

		addChild(m_SIDbin, L"Textual SID: %1!ws!", TextualSid.c_str());
		m_SIDbin->addChild(Revision, L"Revision: 0x%1!02X!", Revision->getData());
		m_SIDbin->addChild(SubAuthorityCount, L"SubAuthorityCount: 0x%1!02X!", SubAuthorityCount->getData());
		if (psia != nullptr)
		{
			const auto is = sid::LookupIdentifierAuthority(*psia);
			m_SIDbin->addChild(IdentifierAuthority, L"IdentifierAuthority: %1!ws!", is.c_str());
		}

		int i = 0;
		for (const auto& sa : SubAuthority)
		{
			m_SIDbin->addChild(sa, L"SubAuthority[%1!d!]: %2!d! = 0x%2!08X!", i, sa->getData());
			i++;
		}
	}
} // namespace smartview