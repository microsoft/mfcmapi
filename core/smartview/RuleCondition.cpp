#include <core/stdafx.h>
#include <core/smartview/RuleCondition.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/interpret/guid.h>
#include <bitset>

namespace smartview
{
	bool PropertyName::isComplete() const noexcept
	{
		if (!Kind->isSet() || !Guid->isSet()) return false;
		if (*Kind == MNID_ID) return LID->isSet();
		if (*Kind == MNID_STRING)
			return NameSize->isSet() && *NameSize % sizeof(WCHAR) == 0 && Name->isSet() && nameTerminated;
		return *Kind == 0xFF;
	}

	void PropertyName::parse()
	{
		// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/32617081-b2fa-4dac-b6c4-0946f51f99d7
		Kind = blockT<BYTE>::parse(parser);
		Guid = blockT<GUID>::parse(parser);
		if (*Kind == MNID_ID)
		{
			LID = blockT<DWORD>::parse(parser);
		}
		else if (*Kind == MNID_STRING)
		{
			NameSize = blockT<BYTE>::parse(parser);
			if (NameSize->isSet() && *NameSize >= sizeof(WCHAR) && parser->checkSize(*NameSize))
			{
				const auto bytes = parser->getAddress();
				nameTerminated = bytes[*NameSize - 2] == 0 && bytes[*NameSize - 1] == 0;
			}

			Name = blockStringW::parse(parser, *NameSize / sizeof(WCHAR));
		}
	}

	void PropertyName::parseBlocks()
	{
		addChild(PropId, L"PropID = 0x%1!04X!", PropId->getData());

		addChild(Kind, L"Kind = 0x%1!02X!", Kind->getData());
		addChild(Guid, L"Guid = %1!ws!", guid::GUIDToString(*Guid).c_str());

		if (*Kind == MNID_ID)
		{
			addChild(LID, L"LID = 0x%1!08X!", LID->getData());
		}
		else if (*Kind == MNID_STRING)
		{
			addChild(NameSize, L"NameSize = 0x%1!02X!", NameSize->getData());
			addChild(Name, L"Name = %1!ws!", Name->c_str());
		}
	}

	void NamedPropertyInformation::parse()
	{
		NoOfNamedProps = blockT<WORD>::parse(parser);
		if (!NoOfNamedProps->isSet() || *NoOfNamedProps >= _MaxEntriesLarge)
		{
			return;
		}

		knownEnd = true;
		if (*NoOfNamedProps)
		{
			// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/487d4ed5-9896-4d72-a512-e612a9d8147f
			std::bitset<0x10000> seen;
			PropId.reserve(*NoOfNamedProps);
			for (auto i = 0; i < *NoOfNamedProps; i++)
			{
				PropId.push_back(blockT<WORD>::parse(parser));
				if (!PropId.back()->isSet())
				{
					knownEnd = false;
					return;
				}

				const auto id = PropId.back()->getData();
				if (id < 0x8000 || seen.test(id)) invalidProperties = true;
				seen.set(id);
			}

			NamedPropertiesSize = blockT<DWORD>::parse(parser);
			if (!NamedPropertiesSize->isSet())
			{
				knownEnd = false;
				return;
			}

			knownEnd = *NamedPropertiesSize <= parser->getSize();
			if (!*NamedPropertiesSize)
			{
				invalidProperties = true;
				return;
			}

			parser->setCap(*NamedPropertiesSize);
			PropertyName.reserve(*NoOfNamedProps);
			for (auto i = 0; i < *NoOfNamedProps; i++)
			{
				auto namedProp = std::make_shared<smartview::PropertyName>(PropId[i]);
				namedProp->block::parse(parser, false);
				if (namedProp->isSet()) PropertyName.emplace_back(namedProp);
				if (!namedProp->isComplete())
				{
					invalidProperties = true;
					break;
				}
			}

			remainingBytes = blockBytes::parse(parser, parser->getSize());
			if (remainingBytes->isSet()) invalidProperties = true;
			parser->clearCap();
		}
	}

	void NamedPropertyInformation::parseBlocks()
	{
		setText(L"NamedPropertyInformation");
		addChild(NoOfNamedProps, L"Number of named props = 0x%1!04X!", NoOfNamedProps->getData());
		addChild(NamedPropertiesSize, L"Named prop size = 0x%1!08X!", NamedPropertiesSize->getData());
		if (!knownEnd) addHeader(L"Truncated or invalid named property information");
		if (invalidProperties) addHeader(L"Invalid named properties");

		for (size_t i = 0; i < PropertyName.size(); i++)
		{
			addChild(PropertyName[i], L"Named Prop 0x%1!04X!", i);
		}

		addLabeledChild(L"Remaining named property bytes", remainingBytes);
	}

	void RuleCondition::parse()
	{
		// A standard PidTagRuleCondition is a bare restriction with no named property prefix:
		// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/fbcda4fc-998c-4de2-b5b3-ccac6fed85fa
		// PidTagExtendedRuleMessageCondition is prefixed by NamedPropertyInformation:
		// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/487d4ed5-9896-4d72-a512-e612a9d8147f
		if (m_bExtended)
		{
			m_NamedPropertyInformation = block::parse<NamedPropertyInformation>(parser, false);
			if (!m_NamedPropertyInformation->hasKnownEnd()) return;
		}

		m_lpRes = std::make_shared<RestrictionStruct>(true, m_bExtended);
		m_lpRes->block::parse(parser, false);
	}

	void RuleCondition::parseBlocks()
	{
		setText(m_bExtended ? L"Extended Rule Condition" : L"Rule Condition");

		addChild(m_NamedPropertyInformation);
		if (m_lpRes && m_lpRes->hasData())
		{
			addChild(m_lpRes);
		}
	}
} // namespace smartview