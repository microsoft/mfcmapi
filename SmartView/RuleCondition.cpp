#include "stdafx.h"
#include "RuleCondition.h"
#include "RestrictionStruct.h"
#include "String.h"
#include "InterpretProp2.h"

RuleCondition::RuleCondition(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, bool bExtended) : SmartViewParser(cbBin, lpBin)
{
	m_NamedPropertyInformation = { 0 };
	m_lpRes = nullptr;
	m_bExtended = bExtended;
}

RuleCondition::~RuleCondition()
{
	delete[] m_NamedPropertyInformation.PropId;
	m_NamedPropertyInformation.NoOfNamedProps;
	if (m_NamedPropertyInformation.PropertyName)
	{
		for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
		{
			delete[] m_NamedPropertyInformation.PropertyName[i].Name;
		}
	}

	delete[] m_NamedPropertyInformation.PropertyName;
	delete m_lpRes;
}

void RuleCondition::Parse()
{
	m_Parser.GetWORD(&m_NamedPropertyInformation.NoOfNamedProps);
	if (m_NamedPropertyInformation.NoOfNamedProps && m_NamedPropertyInformation.NoOfNamedProps < _MaxEntriesLarge)
	{
		m_NamedPropertyInformation.PropId = new WORD[m_NamedPropertyInformation.NoOfNamedProps];
		if (m_NamedPropertyInformation.PropId)
		{
			memset(m_NamedPropertyInformation.PropId, 0, m_NamedPropertyInformation.NoOfNamedProps*sizeof(WORD));
			for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_Parser.GetWORD(&m_NamedPropertyInformation.PropId[i]);
			}
		}

		m_Parser.GetDWORD(&m_NamedPropertyInformation.NamedPropertiesSize);
		m_NamedPropertyInformation.PropertyName = new PropertyNameStruct[m_NamedPropertyInformation.NoOfNamedProps];
		if (m_NamedPropertyInformation.PropertyName)
		{
			memset(m_NamedPropertyInformation.PropertyName, 0, m_NamedPropertyInformation.NoOfNamedProps*sizeof(PropertyNameStruct));
			for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_Parser.GetBYTE(&m_NamedPropertyInformation.PropertyName[i].Kind);
				m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), reinterpret_cast<LPBYTE>(&m_NamedPropertyInformation.PropertyName[i].Guid));
				if (MNID_ID == m_NamedPropertyInformation.PropertyName[i].Kind)
				{
					m_Parser.GetDWORD(&m_NamedPropertyInformation.PropertyName[i].LID);
				}
				else if (MNID_STRING == m_NamedPropertyInformation.PropertyName[i].Kind)
				{
					m_Parser.GetBYTE(&m_NamedPropertyInformation.PropertyName[i].NameSize);
					m_Parser.GetStringW(
						m_NamedPropertyInformation.PropertyName[i].NameSize / sizeof(WCHAR),
						&m_NamedPropertyInformation.PropertyName[i].Name);
				}
			}
		}
	}

	m_lpRes = new RestrictionStruct(
		static_cast<ULONG>(m_Parser.RemainingBytes()),
		m_Parser.GetCurrentAddress(),
		true,
		m_bExtended);
	if (m_lpRes)
	{
		m_lpRes->DisableJunkParsing();
		m_lpRes->EnsureParsed();
		m_Parser.Advance(m_lpRes->GetCurrentOffset());
	}
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

	if (m_NamedPropertyInformation.NoOfNamedProps && m_NamedPropertyInformation.PropId)
	{
		szRuleCondition += formatmessage(IDS_RULECONNAMEPROPSIZE,
			m_NamedPropertyInformation.NamedPropertiesSize);

		for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
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
				if (m_NamedPropertyInformation.PropertyName[i].Name)
				{
					szRuleCondition += m_NamedPropertyInformation.PropertyName[i].Name;
				}
			}
		}
	}

	szRuleCondition += L"\r\n"; // STRING_OK
	szRuleCondition += m_lpRes->ToString();

	return szRuleCondition;
}