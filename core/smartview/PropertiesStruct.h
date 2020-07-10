#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/SPropValueStruct.h>

namespace smartview
{
	class PropertiesStruct : public block
	{
	public:
		PropertiesStruct(DWORD cValues, bool bRuleCondition, bool bNickName)
			: m_MaxEntries(cValues), m_RuleCondition(bRuleCondition), m_NickName(bNickName)
		{
		}
		void SetMaxEntries(DWORD maxEntries) noexcept { m_MaxEntries = maxEntries; }
		void EnableNickNameParsing() noexcept { m_NickName = true; }
		void EnableRuleConditionParsing() noexcept { m_RuleCondition = true; }
		_Check_return_ std::vector<std::shared_ptr<SPropValueStruct>>& Props() noexcept { return m_Props; }

	private:
		void parse() override;
		void parseBlocks() override;

		bool m_NickName{};
		bool m_RuleCondition{};
		DWORD m_MaxEntries{_MaxEntriesSmall};
		std::vector<std::shared_ptr<SPropValueStruct>> m_Props;
	}; // namespace smartview
} // namespace smartview