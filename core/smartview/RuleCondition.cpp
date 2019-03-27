#include <core/stdafx.h>
#include <core/smartview/RuleCondition.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/interpret/guid.h>

namespace smartview
{
	PropertyName::PropertyName(const std::shared_ptr<binaryParser>& parser)
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

	void RuleCondition::Init(bool bExtended) { m_bExtended = bExtended; }

	void RuleCondition::Parse()
	{
		m_NamedPropertyInformation.NoOfNamedProps = blockT<WORD>::parse(m_Parser);
		if (*m_NamedPropertyInformation.NoOfNamedProps && *m_NamedPropertyInformation.NoOfNamedProps < _MaxEntriesLarge)
		{
			m_NamedPropertyInformation.PropId.reserve(*m_NamedPropertyInformation.NoOfNamedProps);
			for (auto i = 0; i < *m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_NamedPropertyInformation.PropId.push_back(blockT<WORD>::parse(m_Parser));
			}

			m_NamedPropertyInformation.NamedPropertiesSize = blockT<DWORD>::parse(m_Parser);

			m_NamedPropertyInformation.PropertyName.reserve(*m_NamedPropertyInformation.NoOfNamedProps);
			for (auto i = 0; i < *m_NamedPropertyInformation.NoOfNamedProps; i++)
			{
				m_NamedPropertyInformation.PropertyName.emplace_back(std::make_shared<PropertyName>(m_Parser));
			}
		}

		m_lpRes = std::make_shared<RestrictionStruct>(true, m_bExtended);
		m_lpRes->smartViewParser::parse(m_Parser, false);
	}

	void RuleCondition::ParseBlocks()
	{
		setRoot(m_bExtended ? L"Extended Rule Condition\r\n" : L"Rule Condition\r\n");

		addChild(
			m_NamedPropertyInformation.NoOfNamedProps,
			L"Number of named props = 0x%1!04X!\r\n",
			m_NamedPropertyInformation.NoOfNamedProps->getData());
		if (!m_NamedPropertyInformation.PropId.empty())
		{
			terminateBlock();
			addChild(
				m_NamedPropertyInformation.NamedPropertiesSize,
				L"Named prop size = 0x%1!08X!",
				m_NamedPropertyInformation.NamedPropertiesSize->getData());

			for (size_t i = 0; i < m_NamedPropertyInformation.PropId.size(); i++)
			{
				terminateBlock();
				addHeader(L"Named Prop 0x%1!04X!\r\n", i);

				addChild(
					m_NamedPropertyInformation.PropId[i],
					L"\tPropID = 0x%1!04X!\r\n",
					m_NamedPropertyInformation.PropId[i]->getData());

				addChild(
					m_NamedPropertyInformation.PropertyName[i]->Kind,
					L"\tKind = 0x%1!02X!\r\n",
					m_NamedPropertyInformation.PropertyName[i]->Kind->getData());
				addChild(
					m_NamedPropertyInformation.PropertyName[i]->Guid,
					L"\tGuid = %1!ws!\r\n",
					guid::GUIDToString(*m_NamedPropertyInformation.PropertyName[i]->Guid).c_str());

				if (m_NamedPropertyInformation.PropertyName[i]->Kind == MNID_ID)
				{
					addChild(
						m_NamedPropertyInformation.PropertyName[i]->LID,
						L"\tLID = 0x%1!08X!",
						m_NamedPropertyInformation.PropertyName[i]->LID->getData());
				}
				else if (*m_NamedPropertyInformation.PropertyName[i]->Kind == MNID_STRING)
				{
					addChild(
						m_NamedPropertyInformation.PropertyName[i]->NameSize,
						L"\tNameSize = 0x%1!02X!\r\n",
						m_NamedPropertyInformation.PropertyName[i]->NameSize->getData());
					addHeader(L"\tName = ");
					addChild(
						m_NamedPropertyInformation.PropertyName[i]->Name,
						m_NamedPropertyInformation.PropertyName[i]->Name->c_str());
				}
			}
		}

		if (m_lpRes && m_lpRes->hasData())
		{
			addChild(m_lpRes->getBlock());
		}
	}
} // namespace smartview