#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockT.h>
#include <core/smartview/block/blockPV.h>

namespace smartview
{
	class SPropValueStruct : public block
	{
	public:
		SPropValueStruct(int index, bool doNickname, bool doRuleProcessing) noexcept
			: m_index(index), m_doNickname(doNickname), m_doRuleProcessing(doRuleProcessing)
		{
		}

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<WORD>> PropType = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> PropID = emptyT<WORD>();
		std::shared_ptr<blockT<ULONG>> ulPropTag = emptyT<ULONG>();
		std::shared_ptr<blockPV> value;
		bool m_doNickname{};
		bool m_doRuleProcessing{};
		int m_index{};
	};
} // namespace smartview