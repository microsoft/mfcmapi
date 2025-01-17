#include <core/stdafx.h>
#include <core/smartview/RuleCondition.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/interpret/guid.h>

namespace smartview
{
	void PropertyName::parse()
	{
		Kind = blockT<BYTE>::parse(parser);
		Guid = blockT<GUID>::parse(parser);
		if (*Kind == MNID_ID)
		{
			LID = blockT<DWORD>::parse(parser);
		}
		else if (*Kind == MNID_STRING)
		{
			NameSize = blockT<BYTE>::parse(parser);
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

	void NamedPropertyInformation ::parse()
	{
		NoOfNamedProps = blockT<WORD>::parse(parser);
		if (*NoOfNamedProps && *NoOfNamedProps < _MaxEntriesLarge)
		{
			PropId.reserve(*NoOfNamedProps);
			for (auto i = 0; i < *NoOfNamedProps; i++)
			{
				PropId.push_back(blockT<WORD>::parse(parser));
			}

			NamedPropertiesSize = blockT<DWORD>::parse(parser);

			PropertyName.reserve(*NoOfNamedProps);
			for (auto i = 0; i < *NoOfNamedProps; i++)
			{
				auto namedProp = std::make_shared<smartview::PropertyName>(PropId[i]);
				namedProp->block::parse(parser, false);
				if (!namedProp->isSet()) break;
				PropertyName.emplace_back(namedProp);
			}
		}
	}

	void NamedPropertyInformation ::parseBlocks()
	{
		setText(L"NamedPropertyInformation");
		addChild(NoOfNamedProps, L"Number of named props = 0x%1!04X!", NoOfNamedProps->getData());
		if (!PropertyName.empty())
		{
			addChild(NamedPropertiesSize, L"Named prop size = 0x%1!08X!", NamedPropertiesSize->getData());

			for (size_t i = 0; i < PropertyName.size(); i++)
			{
				addChild(PropertyName[i], L"Named Prop 0x%1!04X!", i);
			}
		}
	}

	void RuleCondition::parse()
	{
		m_NamedPropertyInformation = block::parse<NamedPropertyInformation>(parser, false);
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