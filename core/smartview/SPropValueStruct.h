#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/SmartView.h>
#include <core/property/parseProperty.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>
#include <core/mapi/mapiFunctions.h>

namespace smartview
{
	class PVBlock : public block
	{
	public:
		PVBlock() = default;
		PVBlock(const PVBlock&) = delete;
		PVBlock& operator=(const PVBlock&) = delete;

		virtual size_t getSize() const noexcept = 0;
		virtual size_t getOffset() const noexcept = 0;
		virtual void getProp(SPropValue& prop) = 0;
		virtual std::wstring propNum(ULONG ulPropTag) { return strings::emptystring; }
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
		ULONG dwAlignPad{};
		std::shared_ptr<PVBlock> value;

		std::shared_ptr<blockT<double>> dbl = emptyT<double>(); /* case PT_DOUBLE */
		std::shared_ptr<blockT<GUID>> lpguid = emptyT<GUID>(); /* case PT_CLSID */
		std::shared_ptr<blockT<LARGE_INTEGER>> li = emptyT<LARGE_INTEGER>(); /* case PT_I8 */
		std::shared_ptr<blockT<SCODE>> err = emptyT<SCODE>(); /* case PT_ERROR */

		_Check_return_ std::shared_ptr<blockStringW> PropBlock()
		{
			EnsurePropBlocks();
			return propBlock;
		}
		_Check_return_ std::shared_ptr<blockStringW> AltPropBlock()
		{
			EnsurePropBlocks();
			return altPropBlock;
		}
		_Check_return_ std::shared_ptr<blockStringW> SmartViewBlock()
		{
			EnsurePropBlocks();
			return smartViewBlock;
		}

		// TODO: Fill in missing cases with test coverage
		void EnsurePropBlocks();

		_Check_return_ std::wstring PropNum() const;

		// Any data we need to cache for getData can live here
	private:
		void parse() override;
		void parseBlocks() override;

		GUID guid{};
		std::shared_ptr<blockStringW> propBlock = emptySW();
		std::shared_ptr<blockStringW> altPropBlock = emptySW();
		std::shared_ptr<blockStringW> smartViewBlock = emptySW();
		bool propStringsGenerated{};
		bool m_doNickname{};
		bool m_doRuleProcessing{};
		int m_index{};
	};
} // namespace smartview