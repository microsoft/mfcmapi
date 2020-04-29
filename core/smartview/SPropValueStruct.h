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
	class FILETIMEBLock;
	struct SBinaryBlock
	{
		std::shared_ptr<blockT<ULONG>> cb = emptyT<ULONG>();
		std::shared_ptr<blockBytes> lpb = emptyBB();
		size_t getSize() const noexcept { return cb->getSize() + lpb->getSize(); }
		size_t getOffset() const noexcept { return cb->getOffset() ? cb->getOffset() : lpb->getOffset(); }

		SBinaryBlock(const std::shared_ptr<binaryParser>& parser)
		{
			cb = blockT<DWORD>::parse(parser);
			// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
			lpb = blockBytes::parse(parser, *cb);
		}
		SBinaryBlock() noexcept {};
	};

	struct SBinaryArrayBlock
	{
		~SBinaryArrayBlock()
		{
			if (binCreated) delete[] bin;
		}
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<SBinaryBlock>> lpbin;
		SBinary* getbin()
		{
			if (binCreated) return bin;
			binCreated = true;
			if (cValues)
			{
				auto count = cValues->getData();
				bin = new (std::nothrow) SBinary[count];
				if (bin)
				{
					for (ULONG i = 0; i < count; i++)
					{
						bin[i].cb = lpbin[i]->cb->getData();
						bin[i].lpb = const_cast<BYTE*>(lpbin[i]->lpb->data());
					}
				}
			}

			return bin;
		}

	private:
		SBinary* bin{};
		bool binCreated{false};
	};

	struct CountedStringA
	{
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
		std::shared_ptr<blockStringA> str = emptySA();
		size_t getSize() const noexcept { return cb->getSize() + str->getSize(); }
		size_t getOffset() const noexcept { return cb->getOffset() ? cb->getOffset() : str->getOffset(); }
	};

	struct CountedStringW
	{
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
		std::shared_ptr<blockStringW> str = emptySW();
		size_t getSize() const noexcept { return cb->getSize() + str->getSize(); }
		size_t getOffset() const noexcept { return cb->getOffset() ? cb->getOffset() : str->getOffset(); }
	};

	struct StringArrayA
	{
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringA>> lppszA;
	};

	struct StringArrayW
	{
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringW>> lppszW;
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
		std::shared_ptr<blockT<WORD>> i = emptyT<WORD>(); /* case PT_I2 */
		std::shared_ptr<blockT<LONG>> l = emptyT<LONG>(); /* case PT_LONG */
		std::shared_ptr<blockT<WORD>> b = emptyT<WORD>(); /* case PT_BOOLEAN */
		std::shared_ptr<blockT<float>> flt = emptyT<float>(); /* case PT_R4 */
		std::shared_ptr<blockT<double>> dbl = emptyT<double>(); /* case PT_DOUBLE */
		std::shared_ptr<FILETIMEBLock> ft; /* case PT_SYSTIME */
		CountedStringA lpszA; /* case PT_STRING8 */
		SBinaryBlock bin; /* case PT_BINARY */
		CountedStringW lpszW; /* case PT_UNICODE */
		std::shared_ptr<blockT<GUID>> lpguid = emptyT<GUID>(); /* case PT_CLSID */
		std::shared_ptr<blockT<LARGE_INTEGER>> li = emptyT<LARGE_INTEGER>(); /* case PT_I8 */
		SBinaryArrayBlock MVbin; /* case PT_MV_BINARY */
		StringArrayA MVszA; /* case PT_MV_STRING8 */
		StringArrayW MVszW; /* case PT_MV_UNICODE */
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