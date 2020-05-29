#pragma once
#include <core/smartview/block/smartViewParser.h>

namespace smartview
{
	struct SPropValueStruct;
	class PropertiesStruct : public smartViewParser
	{
	public:
		void parse(const std::shared_ptr<binaryParser>& parser, DWORD cValues, bool bRuleCondition);
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
	};
} // namespace smartview