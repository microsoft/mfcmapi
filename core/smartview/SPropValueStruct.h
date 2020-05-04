#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/SmartView.h>
#include <core/property/parseProperty.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class blockPV : public block
	{
	public:
		blockPV() = default;
		void parse(std::shared_ptr<binaryParser>& parser, ULONG ulPropTag, bool doNickname, bool doRuleProcessing)

		{
			m_Parser = parser;
			m_doNickname = doNickname;
			m_doRuleProcessing = doRuleProcessing;
			m_ulPropTag = ulPropTag;

			// Offset will always be where we start parsing
			setOffset(m_Parser->getOffset());
			parse();
			// And size will always be how many bytes we consumed
			setSize(m_Parser->getOffset() - getOffset());
		}
		blockPV(const blockPV&) = delete;
		blockPV& operator=(const blockPV&) = delete;

		virtual std::wstring toNumberAsString() { return strings::emptystring; }

		_Check_return_ std::shared_ptr<blockStringW> PropBlock()
		{
			ensurePropBlocks();
			return propBlock;
		}
		_Check_return_ std::shared_ptr<blockStringW> AltPropBlock()
		{
			ensurePropBlocks();
			return altPropBlock;
		}
		_Check_return_ std::shared_ptr<blockStringW> SmartViewBlock()
		{
			ensurePropBlocks();
			return smartViewBlock;
		}

	protected:
		bool m_doNickname{};
		bool m_doRuleProcessing{};
		std::shared_ptr<binaryParser> m_Parser{};
		ULONG m_ulPropTag{};

	private:
		void ensurePropBlocks()
		{
			if (propStringsGenerated) return;
			auto prop = SPropValue{};
			prop.ulPropTag = m_ulPropTag;
			getProp(prop);

			auto propString = std::wstring{};
			auto altPropString = std::wstring{};
			property::parseProperty(&prop, &propString, &altPropString);

			propBlock = std::make_shared<blockStringW>(
				strings::RemoveInvalidCharactersW(propString, false), getSize(), getOffset());

			altPropBlock = std::make_shared<blockStringW>(
				strings::RemoveInvalidCharactersW(altPropString, false), getSize(), getOffset());

			const auto smartViewString = parsePropertySmartView(&prop, nullptr, nullptr, nullptr, false, false);
			smartViewBlock = std::make_shared<blockStringW>(smartViewString, getSize(), getOffset());

			propStringsGenerated = true;
		}

		virtual void parse() = 0;
		virtual void getProp(SPropValue& prop) = 0;
		std::shared_ptr<blockStringW> propBlock = emptySW();
		std::shared_ptr<blockStringW> altPropBlock = emptySW();
		std::shared_ptr<blockStringW> smartViewBlock = emptySW();
		bool propStringsGenerated{};
	};

	struct SPropValueStruct : public smartViewParser
	{
	public:
		void Init(int index, bool doNickname, bool doRuleProcessing) noexcept
		{
			m_index = index;
			m_doNickname = doNickname;
			m_doRuleProcessing = doRuleProcessing;
		}

		std::shared_ptr<blockT<WORD>> PropType = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> PropID = emptyT<WORD>();
		std::shared_ptr<blockT<ULONG>> ulPropTag = emptyT<ULONG>();
		std::shared_ptr<blockPV> value;

	private:
		void parse() override;
		void parseBlocks() override;

		bool m_doNickname{};
		bool m_doRuleProcessing{};
		int m_index{};
	};
} // namespace smartview