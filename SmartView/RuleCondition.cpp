#include "stdafx.h"
#include "..\stdafx.h"
#include "RuleCondition.h"
#include "RestrictionStruct.h"
#include "..\String.h"
#include "..\ParseProperty.h"
#include "..\InterpretProp.h"

RuleCondition::RuleCondition(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, bool bExtended) : SmartViewParser(cbBin, lpBin)
{
	m_NamedPropertyInformation = { 0 };
	m_lpRes = NULL;
	m_bExtended = bExtended;
}

RuleCondition::~RuleCondition()
{
	delete[] m_NamedPropertyInformation.PropId;
	WORD i = 0;
	m_NamedPropertyInformation.NoOfNamedProps;
	if (m_NamedPropertyInformation.PropertyName)
	{
		for (i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
		{
			delete[] m_NamedPropertyInformation.PropertyName[i].Name;
		}
	}

	delete[] m_NamedPropertyInformation.PropertyName;
	DeleteRestriction(m_lpRes);
	delete[] m_lpRes;
}

void RuleCondition::Parse()
{
	WORD i = 0;

	m_Parser.GetWORD(&m_NamedPropertyInformation.NoOfNamedProps);
	if (m_NamedPropertyInformation.NoOfNamedProps && m_NamedPropertyInformation.NoOfNamedProps < _MaxEntriesLarge)
	{
		m_NamedPropertyInformation.PropId = new WORD[m_NamedPropertyInformation.NoOfNamedProps];
		if (m_NamedPropertyInformation.PropId)
		{
			memset(m_NamedPropertyInformation.PropId, 0, m_NamedPropertyInformation.NoOfNamedProps*sizeof(WORD));
			for (i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_Parser.GetWORD(&m_NamedPropertyInformation.PropId[i]);
			}
		}

		m_Parser.GetDWORD(&m_NamedPropertyInformation.NamedPropertiesSize);
		m_NamedPropertyInformation.PropertyName = new PropertyNameStruct[m_NamedPropertyInformation.NoOfNamedProps];
		if (m_NamedPropertyInformation.PropertyName)
		{
			memset(m_NamedPropertyInformation.PropertyName, 0, m_NamedPropertyInformation.NoOfNamedProps*sizeof(PropertyNameStruct));
			for (i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_Parser.GetBYTE(&m_NamedPropertyInformation.PropertyName[i].Kind);
				m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), (LPBYTE)&m_NamedPropertyInformation.PropertyName[i].Guid);
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

	m_lpRes = new SRestriction;
	if (m_lpRes)
	{
		memset(m_lpRes, 0, sizeof(SRestriction));
		size_t cbBytesRead = 0;

		(void)BinToRestriction(
			0,
			(ULONG)m_Parser.RemainingBytes(),
			m_Parser.GetCurrentAddress(),
			&cbBytesRead,
			m_lpRes,
			true,
			m_bExtended);
		m_Parser.Advance(cbBytesRead);
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

		WORD i = 0;
		for (i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
		{
			szRuleCondition += formatmessage(IDS_RULECONNAMEPROPID, i, m_NamedPropertyInformation.PropId[i]);

			szRuleCondition += formatmessage(IDS_RULECONNAMEPROPKIND,
				m_NamedPropertyInformation.PropertyName[i].Kind);

			LPTSTR szGUID = GUIDToString(&m_NamedPropertyInformation.PropertyName[i].Guid);
			szRuleCondition += LPTSTRToWstring(szGUID);
			delete[] szGUID;
			szGUID = NULL;

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
	wstring szRestriction = formatmessage(IDS_RESTRICTIONDATA);
	szRestriction += RestrictionToWstring(m_lpRes, NULL);
	szRuleCondition += szRestriction;

	return szRuleCondition;
}