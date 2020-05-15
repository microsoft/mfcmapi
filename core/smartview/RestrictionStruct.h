#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class RestrictionStruct;
	class blockRes : public block
	{
	public:
		blockRes() = default;
		void parse(std::shared_ptr<binaryParser>& parser, ULONG ulDepth, bool bRuleCondition, bool bExtendedCount)
		{
			m_Parser = parser;
			m_ulDepth = ulDepth;
			m_bRuleCondition = bRuleCondition;
			m_bExtendedCount = bExtendedCount;

			// Offset will always be where we start parsing
			setOffset(m_Parser->getOffset());
			parse();
			// And size will always be how many bytes we consumed
			setSize(m_Parser->getOffset() - getOffset());
		}
		blockRes(const blockRes&) = delete;
		blockRes& operator=(const blockRes&) = delete;
		virtual void parseBlocks(ULONG ulTabLevel) = 0;

	protected:
		const inline std::wstring makeTabs(ULONG ulTabLevel) const { return std::wstring(ulTabLevel, L'\t'); }
		ULONG m_ulDepth{};
		bool m_bRuleCondition{};
		bool m_bExtendedCount{};
		std::shared_ptr<binaryParser> m_Parser{};

	private:
		virtual void parse() = 0;
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
		}
		void parseBlocks(ULONG ulTabLevel);

	private:
		void parse() override { parse(0); }
		void parse(ULONG ulDepth);
		void parseBlocks() override
		{
			setRoot(L"Restriction:\r\n");
			parseBlocks(0);
		};

		std::shared_ptr<blockT<DWORD>> rt = emptyT<DWORD>(); /* Restriction type */
		std::shared_ptr<blockRes> res1; // TODO: fix name

		bool m_bRuleCondition{};
		bool m_bExtendedCount{};
	};
} // namespace smartview