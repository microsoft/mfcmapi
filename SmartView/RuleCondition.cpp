#include "stdafx.h"
#include "RuleCondition.h"
#include "RestrictionStruct.h"
#include "String.h"
#include "InterpretProp2.h"

RuleCondition::RuleCondition()
{
	m_bExtended = false;
}

void RuleCondition::Init(bool bExtended)
{
	m_bExtended = bExtended;
}

void RuleCondition::Parse()
{
	m_Parser.GetWORD(&m_NamedPropertyInformation.NoOfNamedProps);
	if (m_NamedPropertyInformation.NoOfNamedProps && m_NamedPropertyInformation.NoOfNamedProps < _MaxEntriesLarge)
	{
		{
			for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				WORD propId = 0;
				m_Parser.GetWORD(&propId);
				m_NamedPropertyInformation.PropId.push_back(propId);
			}
		}

		m_Parser.GetDWORD(&m_NamedPropertyInformation.NamedPropertiesSize);
		{
			for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				PropertyName propertyName;
				m_Parser.GetBYTE(&propertyName.Kind);
				propertyName.Guid = m_Parser.Get<GUID>();
				if (MNID_ID == propertyName.Kind)
				{
					m_Parser.GetDWORD(&propertyName.LID);
				}
				else if (MNID_STRING == propertyName.Kind)
				{
					m_Parser.GetBYTE(&propertyName.NameSize);
					propertyName.Name = m_Parser.GetStringW(propertyName.NameSize / sizeof(WCHAR));
				}

				m_NamedPropertyInformation.PropertyName.push_back(propertyName);
			}
		}
	}

	m_lpRes.Init(true, m_bExtended);
	m_lpRes.SmartViewParser::Init(
		static_cast<ULONG>(m_Parser.RemainingBytes()),
		m_Parser.GetCurrentAddress());
	m_lpRes.DisableJunkParsing();
	m_lpRes.EnsureParsed();
	m_Parser.Advance(m_lpRes.GetCurrentOffset());
}

_Check_return_ wstring RuleCondition::ToStringInternal()
{
	wstring szRuleCondition;

	if (m_bExtended)
	{
		szRuleCondition = formatmessage(IDS_EXRULECONHEADER,
			m_NamedPropertyInformation.NoOfNamedProps);
	}
	else
	{
		szRuleCondition = formatmessage(IDS_RULECONHEADER,
			m_NamedPropertyInformation.NoOfNamedProps);
	}

	if (m_NamedPropertyInformation.PropId.size())
	{
		szRuleCondition += formatmessage(IDS_RULECONNAMEPROPSIZE,
			m_NamedPropertyInformation.NamedPropertiesSize);

		for (size_t i = 0; i < m_NamedPropertyInformation.PropId.size(); i++)
		{
			szRuleCondition += formatmessage(IDS_RULECONNAMEPROPID, i, m_NamedPropertyInformation.PropId[i]);

			szRuleCondition += formatmessage(IDS_RULECONNAMEPROPKIND,
				m_NamedPropertyInformation.PropertyName[i].Kind);

			szRuleCondition += GUIDToString(&m_NamedPropertyInformation.PropertyName[i].Guid);

			if (MNID_ID == m_NamedPropertyInformation.PropertyName[i].Kind)
			{
				szRuleCondition += formatmessage(IDS_RULECONNAMEPROPLID,
					m_NamedPropertyInformation.PropertyName[i].LID);
			}
			else if (MNID_STRING == m_NamedPropertyInformation.PropertyName[i].Kind)
			{
				szRuleCondition += formatmessage(IDS_RULENAMEPROPSIZE,
					m_NamedPropertyInformation.PropertyName[i].NameSize);
				szRuleCondition += m_NamedPropertyInformation.PropertyName[i].Name;
			}
		}
	}

	szRuleCondition += L"\r\n"; // STRING_OK
	szRuleCondition += m_lpRes.ToString();

	return szRuleCondition;
}