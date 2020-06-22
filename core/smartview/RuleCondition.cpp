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

	void RuleCondition::Init(bool bExtended) noexcept { m_bExtended = bExtended; }

	void RuleCondition::parse()
	{
		m_NamedPropertyInformation.NoOfNamedProps = blockT<WORD>::parse(parser);
		if (*m_NamedPropertyInformation.NoOfNamedProps && *m_NamedPropertyInformation.NoOfNamedProps < _MaxEntriesLarge)
		{
			m_NamedPropertyInformation.PropId.reserve(*m_NamedPropertyInformation.NoOfNamedProps);
			for (auto i = 0; i < *m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_NamedPropertyInformation.PropId.push_back(blockT<WORD>::parse(parser));
			}

			m_NamedPropertyInformation.NamedPropertiesSize = blockT<DWORD>::parse(parser);

			m_NamedPropertyInformation.PropertyName.reserve(*m_NamedPropertyInformation.NoOfNamedProps);
			for (auto i = 0; i < *m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				auto namedProp = block::parse<PropertyName>(parser, false);
				m_NamedPropertyInformation.PropertyName.emplace_back(namedProp);
			}
		}

		m_lpRes = std::make_shared<RestrictionStruct>(true, m_bExtended);
		m_lpRes->block::parse(parser, false);
	}

	void RuleCondition::parseBlocks()
	{
		setText(m_bExtended ? L"Extended Rule Condition" : L"Rule Condition");

		addChild(
			m_NamedPropertyInformation.NoOfNamedProps,
			L"Number of named props = 0x%1!04X!",
			m_NamedPropertyInformation.NoOfNamedProps->getData());
		if (!m_NamedPropertyInformation.PropId.empty())
		{
			addChild(
				m_NamedPropertyInformation.NamedPropertiesSize,
				L"Named prop size = 0x%1!08X!",
				m_NamedPropertyInformation.NamedPropertiesSize->getData());

			for (size_t i = 0; i < m_NamedPropertyInformation.PropId.size(); i++)
			{
				auto namedProp = m_NamedPropertyInformation.PropertyName[i];
				addChild(namedProp);
				namedProp->setText(L"Named Prop 0x%1!04X!", i);

				namedProp->addChild(
					m_NamedPropertyInformation.PropId[i],
					L"PropID = 0x%1!04X!",
					m_NamedPropertyInformation.PropId[i]->getData());

				namedProp->addChild(
					m_NamedPropertyInformation.PropertyName[i]->Kind,
					L"Kind = 0x%1!02X!",
					m_NamedPropertyInformation.PropertyName[i]->Kind->getData());
				namedProp->addChild(
					m_NamedPropertyInformation.PropertyName[i]->Guid,
					L"Guid = %1!ws!",
					guid::GUIDToString(*m_NamedPropertyInformation.PropertyName[i]->Guid).c_str());

				if (*m_NamedPropertyInformation.PropertyName[i]->Kind == MNID_ID)
				{
					namedProp->addChild(
						m_NamedPropertyInformation.PropertyName[i]->LID,
						L"LID = 0x%1!08X!",
						m_NamedPropertyInformation.PropertyName[i]->LID->getData());
				}
				else if (*m_NamedPropertyInformation.PropertyName[i]->Kind == MNID_STRING)
				{
					namedProp->addChild(
						m_NamedPropertyInformation.PropertyName[i]->NameSize,
						L"NameSize = 0x%1!02X!",
						m_NamedPropertyInformation.PropertyName[i]->NameSize->getData());
					namedProp->addChild(
						m_NamedPropertyInformation.PropertyName[i]->Name,
						L"Name = %1!ws!",
						m_NamedPropertyInformation.PropertyName[i]->Name->c_str());
				}
			}
		}

		if (m_lpRes && m_lpRes->hasData())
		{
			addChild(m_lpRes);
		}
	}
} // namespace smartview