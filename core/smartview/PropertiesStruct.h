#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/SPropValueStruct.h>

namespace smartview
{
	class PropertiesStruct : public block
	{
	public:
		PropertiesStruct(DWORD cValues, bool bNickName, bool bRuleCondition)
			: m_MaxEntries(cValues), m_NickName(bNickName), m_RuleCondition(bRuleCondition)
		{
		}

	private:
		void parse() override;
		void parseBlocks() override;

		bool m_NickName{};
		bool m_RuleCondition{};
		DWORD m_MaxEntries{_MaxEntriesSmall};
		std::vector<std::shared_ptr<SPropValueStruct>> m_Props;
	};
} // namespace smartview