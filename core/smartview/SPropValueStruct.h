#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/SmartView.h>
#include <core/property/parseProperty.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class PVBlock : public block
	{
	public:
		PVBlock() = default;
		void parse(std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)

		{
			m_Parser = parser;
			m_doNickname = doNickname;
			m_doRuleProcessing = doRuleProcessing;
			parse();
		}
		PVBlock(const PVBlock&) = delete;
		PVBlock& operator=(const PVBlock&) = delete;


		virtual std::wstring propNum(ULONG /*ulPropTag*/) { return strings::emptystring; }

		_Check_return_ std::shared_ptr<blockStringW> PropBlock(ULONG ulPropTag)
		{
			ensurePropBlocks(ulPropTag);
			return propBlock;
		}
		_Check_return_ std::shared_ptr<blockStringW> AltPropBlock(ULONG ulPropTag)
		{
			ensurePropBlocks(ulPropTag);
			return altPropBlock;
		}
		_Check_return_ std::shared_ptr<blockStringW> SmartViewBlock(ULONG ulPropTag)
		{
			ensurePropBlocks(ulPropTag);
			return smartViewBlock;
		}

	protected:
		bool m_doNickname{};
		bool m_doRuleProcessing{};
		std::shared_ptr<binaryParser> m_Parser{};

	private:
		void ensurePropBlocks(ULONG ulPropTag)
		{
			if (propStringsGenerated) return;
			auto prop = SPropValue{};
			prop.ulPropTag = ulPropTag;
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
		std::shared_ptr<PVBlock> value;

		_Check_return_ std::wstring PropNum() const;

	private:
		void parse() override;
		void parseBlocks() override;

		bool m_doNickname{};
		bool m_doRuleProcessing{};
		int m_index{};
	};
} // namespace smartview