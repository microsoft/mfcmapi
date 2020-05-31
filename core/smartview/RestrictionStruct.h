#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class blockRes : public smartViewParser
	{
	public:
		blockRes() = default;
		void parse(std::shared_ptr<binaryParser>& parser, ULONG ulDepth, bool bRuleCondition, bool bExtendedCount)
		{
			m_Parser = parser;
			m_ulDepth = ulDepth;
			m_bRuleCondition = bRuleCondition;
			m_bExtendedCount = bExtendedCount;

			ensureParsed();
		}
		blockRes(const blockRes&) = delete;
		blockRes& operator=(const blockRes&) = delete;
		virtual void parseBlocks(ULONG ulTabLevel) = 0;

	protected:
		const inline std::wstring makeTabs(ULONG ulTabLevel) const { return std::wstring(ulTabLevel, L'\t'); }
		ULONG m_ulDepth{};
		bool m_bRuleCondition{};
		bool m_bExtendedCount{};

	private:
		void parse() override = 0;
		void parseBlocks() override{};
	};

	class RestrictionStruct : public smartViewParser
	{
	public:
		RestrictionStruct(bool bRuleCondition, bool bExtendedCount)
			: m_bRuleCondition(bRuleCondition), m_bExtendedCount(bExtendedCount)
		{
		}
		RestrictionStruct(
			const std::shared_ptr<binaryParser>& parser,
			ULONG ulDepth,
			bool bRuleCondition,
			bool bExtendedCount)
			: m_bRuleCondition(bRuleCondition), m_bExtendedCount(bExtendedCount)
		{
			m_Parser = parser;
			parse(ulDepth);
			parsed = true;
		}
		void parseBlocks(ULONG ulTabLevel);

	private:
		void parse() override { parse(0); }
		void parse(ULONG ulDepth);
		void parseBlocks() override
		{
			setText(L"Restriction:\r\n");
			parseBlocks(0);
		};

		std::shared_ptr<blockT<DWORD>> rt = emptyT<DWORD>(); /* Restriction type */
		std::shared_ptr<blockRes> res;

		bool m_bRuleCondition{};
		bool m_bExtendedCount{};
	};
} // namespace smartview