#include <core/stdafx.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	void PropertiesStruct::parse(const std::shared_ptr<binaryParser>& parser, DWORD cValues, bool bRuleCondition)
	{
		SetMaxEntries(cValues);
		if (bRuleCondition) EnableRuleConditionParsing();
		smartViewParser::parse(parser, false);
	}

	void PropertiesStruct::parse()
	{
		// For consistancy with previous parsings, we'll refuse to parse if asked to parse more than _MaxEntriesSmall
		// However, we may want to reconsider this choice.
		if (m_MaxEntries > _MaxEntriesSmall) return;

		DWORD dwPropCount = 0;

		// If we have a non-default max, it was computed elsewhere and we do expect to have that many entries. So we can reserve.
		if (m_MaxEntries != _MaxEntriesSmall)
		{
			m_Props.reserve(m_MaxEntries);
		}

		for (;;)
		{
			if (dwPropCount >= m_MaxEntries) break;
			auto parser = std::make_shared<SPropValueStruct>();
			if (parser)
			{
				parser->Init(dwPropCount++, m_NickName, m_RuleCondition);
				parser->smartViewParser::parse(m_Parser, 0, false);
			}

			m_Props.emplace_back(parser);
		}
	}

	void PropertiesStruct::parseBlocks()
	{
		for (const auto& prop : m_Props)
		{
			addChild(prop->getBlock());
			terminateBlock();
		}
	}
} // namespace smartview