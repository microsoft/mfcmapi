#include <StdAfx.h>
#include <Interpret/SmartView/RuleCondition.h>
#include <Interpret/SmartView/RestrictionStruct.h>
#include <Interpret/String.h>
#include <Interpret/Guids.h>

namespace smartview
{
	void RuleCondition::Init(bool bExtended) { m_bExtended = bExtended; }

	void RuleCondition::Parse()
	{
		m_NamedPropertyInformation.NoOfNamedProps = m_Parser.Get<WORD>();
		if (m_NamedPropertyInformation.NoOfNamedProps && m_NamedPropertyInformation.NoOfNamedProps < _MaxEntriesLarge)
		{
			m_NamedPropertyInformation.PropId.reserve(m_NamedPropertyInformation.NoOfNamedProps);
			for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_NamedPropertyInformation.PropId.push_back(m_Parser.Get<WORD>());
			}

			m_NamedPropertyInformation.NamedPropertiesSize = m_Parser.Get<DWORD>();

			m_NamedPropertyInformation.PropertyName.reserve(m_NamedPropertyInformation.NoOfNamedProps);
			for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				PropertyName propertyName;
				propertyName.Kind = m_Parser.Get<BYTE>();
				propertyName.Guid = m_Parser.Get<GUID>();
				if (propertyName.Kind == MNID_ID)
				{
					propertyName.LID = m_Parser.Get<DWORD>();
				}
				else if (propertyName.Kind == MNID_STRING)
				{
					propertyName.NameSize = m_Parser.Get<BYTE>();
					propertyName.Name = m_Parser.GetStringW(propertyName.NameSize / sizeof(WCHAR));
				}

				m_NamedPropertyInformation.PropertyName.push_back(propertyName);
			}
		}

		m_lpRes.init(true, m_bExtended);
		m_lpRes.SmartViewParser::init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
		m_lpRes.DisableJunkParsing();
		m_lpRes.EnsureParsed();
		m_Parser.advance(m_lpRes.GetCurrentOffset());
	}

	void RuleCondition::ParseBlocks()
	{
		addHeader(m_bExtended ? L"Extended Rule Condition" : L"Rule Condition");

		addLine();
		addBlock(
			m_NamedPropertyInformation.NoOfNamedProps,
			L"Number of named props = 0x%1!04X!",
			m_NamedPropertyInformation.NoOfNamedProps.getData());
		addLine();
		if (!m_NamedPropertyInformation.PropId.empty())
		{
			addLine();
			addBlock(
				m_NamedPropertyInformation.NamedPropertiesSize,
				L"Named prop size = 0x%1!08X!",
				m_NamedPropertyInformation.NamedPropertiesSize.getData());

			for (size_t i = 0; i < m_NamedPropertyInformation.PropId.size(); i++)
			{
				addLine();
				addHeader(L"Named Prop 0x%1!04X!\r\n", i);
				addBlock(
					m_NamedPropertyInformation.PropId[i],
					L"\tPropID = 0x%2!04X!\r\n",
					m_NamedPropertyInformation.PropId[i].getData());

				addBlock(
					m_NamedPropertyInformation.PropertyName[i].Kind,
					L"\tKind = 0x%1!02X!\r\n",
					m_NamedPropertyInformation.PropertyName[i].Kind.getData());
				addBlock(
					m_NamedPropertyInformation.PropertyName[i].Guid,
					L"\tGuid = %1!ws!",
					guid::GUIDToString(m_NamedPropertyInformation.PropertyName[i].Guid).c_str());

				addLine();
				if (m_NamedPropertyInformation.PropertyName[i].Kind == MNID_ID)
				{
					addBlock(
						m_NamedPropertyInformation.PropertyName[i].LID,
						L"\tLID = 0x%1!08X!",
						m_NamedPropertyInformation.PropertyName[i].LID.getData());
				}
				else if (m_NamedPropertyInformation.PropertyName[i].Kind == MNID_STRING)
				{
					addBlock(
						m_NamedPropertyInformation.PropertyName[i].NameSize,
						L"\tNameSize = 0x%1!02X!",
						m_NamedPropertyInformation.PropertyName[i].NameSize.getData());
					addLine();
					addHeader(L"\tName = ");
					addBlock(
						m_NamedPropertyInformation.PropertyName[i].Name,
						m_NamedPropertyInformation.PropertyName[i].Name.c_str());
				}
			}
		}

		addBlock(m_lpRes.getBlock());
	}
} // namespace smartview