#include <core/stdafx.h>
#include <core/smartview/RuleCondition.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/interpret/guid.h>

namespace smartview
{
	PropertyName::PropertyName(std::shared_ptr<binaryParser>& parser)
	{
		Kind = parser->Get<BYTE>();
		Guid = parser->Get<GUID>();
		if (Kind == MNID_ID)
		{
			LID = parser->Get<DWORD>();
		}
		else if (Kind == MNID_STRING)
		{
			NameSize = parser->Get<BYTE>();
			Name.parse(parser, NameSize / sizeof(WCHAR));
		}
	}

	void RuleCondition::Init(bool bExtended) { m_bExtended = bExtended; }

	void RuleCondition::Parse()
	{
		m_NamedPropertyInformation.NoOfNamedProps = m_Parser->Get<WORD>();
		if (m_NamedPropertyInformation.NoOfNamedProps && m_NamedPropertyInformation.NoOfNamedProps < _MaxEntriesLarge)
		{
			m_NamedPropertyInformation.PropId.reserve(m_NamedPropertyInformation.NoOfNamedProps);
			for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_NamedPropertyInformation.PropId.push_back(m_Parser->Get<WORD>());
			}

			m_NamedPropertyInformation.NamedPropertiesSize = m_Parser->Get<DWORD>();

			m_NamedPropertyInformation.PropertyName.reserve(m_NamedPropertyInformation.NoOfNamedProps);
			for (auto i = 0; i < m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_NamedPropertyInformation.PropertyName.emplace_back(std::make_shared<PropertyName>(m_Parser));
			}
		}

		m_lpRes = std::make_shared<RestrictionStruct>(true, m_bExtended);
		m_lpRes->SmartViewParser::parse(m_Parser, false);
	}

	void RuleCondition::ParseBlocks()
	{
		setRoot(m_bExtended ? L"Extended Rule Condition\r\n" : L"Rule Condition\r\n");

		addBlock(
			m_NamedPropertyInformation.NoOfNamedProps,
			L"Number of named props = 0x%1!04X!\r\n",
			m_NamedPropertyInformation.NoOfNamedProps.getData());
		if (!m_NamedPropertyInformation.PropId.empty())
		{
			terminateBlock();
			addBlock(
				m_NamedPropertyInformation.NamedPropertiesSize,
				L"Named prop size = 0x%1!08X!",
				m_NamedPropertyInformation.NamedPropertiesSize.getData());

			for (size_t i = 0; i < m_NamedPropertyInformation.PropId.size(); i++)
			{
				terminateBlock();
				addHeader(L"Named Prop 0x%1!04X!\r\n", i);
				addBlock(
					m_NamedPropertyInformation.PropId[i],
					L"\tPropID = 0x%1!04X!\r\n",
					m_NamedPropertyInformation.PropId[i].getData());

				addBlock(
					m_NamedPropertyInformation.PropertyName[i]->Kind,
					L"\tKind = 0x%1!02X!\r\n",
					m_NamedPropertyInformation.PropertyName[i]->Kind.getData());
				addBlock(
					m_NamedPropertyInformation.PropertyName[i]->Guid,
					L"\tGuid = %1!ws!\r\n",
					guid::GUIDToString(m_NamedPropertyInformation.PropertyName[i]->Guid).c_str());

				if (m_NamedPropertyInformation.PropertyName[i]->Kind == MNID_ID)
				{
					addBlock(
						m_NamedPropertyInformation.PropertyName[i]->LID,
						L"\tLID = 0x%1!08X!",
						m_NamedPropertyInformation.PropertyName[i]->LID.getData());
				}
				else if (m_NamedPropertyInformation.PropertyName[i]->Kind == MNID_STRING)
				{
					addBlock(
						m_NamedPropertyInformation.PropertyName[i]->NameSize,
						L"\tNameSize = 0x%1!02X!\r\n",
						m_NamedPropertyInformation.PropertyName[i]->NameSize.getData());
					addHeader(L"\tName = ");
					addBlock(
						m_NamedPropertyInformation.PropertyName[i]->Name,
						m_NamedPropertyInformation.PropertyName[i]->Name.c_str());
				}
			}
		}

		if (m_lpRes && m_lpRes->hasData())
		{
			addBlock(m_lpRes->getBlock());
		}
	}
} // namespace smartview