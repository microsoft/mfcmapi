#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class blockRes : public block
	{
	public:
		void init(ULONG ulDepth, bool bRuleCondition, bool bExtendedCount)
		{
			m_ulDepth = ulDepth;
			m_bRuleCondition = bRuleCondition;
			m_bExtendedCount = bExtendedCount;
		}

	protected:
		const inline std::wstring makeTabs(ULONG ulTabLevel) const { return std::wstring(ulTabLevel, L'\t'); }
		ULONG m_ulDepth{};
		bool m_bRuleCondition{};
		bool m_bExtendedCount{};

	private:
		void parse() override = 0;
		void parseBlocks() override = 0;
		bool usePipes() const override { return true; }
	};

	class RestrictionStruct : public block
	{
	public:
		RestrictionStruct(ULONG ulDepth, bool bRuleCondition, bool bExtendedCount)
			: m_ulDepth(ulDepth), m_bRuleCondition(bRuleCondition), m_bExtendedCount(bExtendedCount)
		{
		}
		RestrictionStruct(bool bRuleCondition, bool bExtendedCount)
			: m_ulDepth(0), m_bRuleCondition(bRuleCondition), m_bExtendedCount(bExtendedCount)
		{
		}

	private:
		void parse() override;
		void parseBlocks() override;
		bool usePipes() const override { return true; }

		std::shared_ptr<blockT<DWORD>> rt = emptyT<DWORD>(); /* Restriction type */
		std::shared_ptr<blockRes> res;

		bool m_bRuleCondition{};
		bool m_bExtendedCount{};
		ULONG m_ulDepth{};
	};
} // namespace smartview