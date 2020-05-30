#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/smartview/SmartView.h>
#include <core/property/parseProperty.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
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
			auto prop = SPropValue{m_ulPropTag, 0, {}};
			getProp(prop);

			auto propString = std::wstring{};
			auto altPropString = std::wstring{};
			property::parseProperty(&prop, &propString, &altPropString);

			propBlock =
				blockStringW::parse(strings::RemoveInvalidCharactersW(propString, false), getSize(), getOffset());

			altPropBlock =
				blockStringW::parse(strings::RemoveInvalidCharactersW(altPropString, false), getSize(), getOffset());

			const auto smartViewString = parsePropertySmartView(&prop, nullptr, nullptr, nullptr, false, false);
			smartViewBlock = blockStringW::parse(smartViewString, getSize(), getOffset());

			propStringsGenerated = true;
		}

		virtual void parse() = 0;
		virtual const void getProp(SPropValue& prop) noexcept = 0;
		std::shared_ptr<blockStringW> propBlock = emptySW();
		std::shared_ptr<blockStringW> altPropBlock = emptySW();
		std::shared_ptr<blockStringW> smartViewBlock = emptySW();
		bool propStringsGenerated{};
	};

	/* case PT_SYSTIME */
	class FILETIMEBLock : public blockPV
	{
	private:
		void parse() override
		{
			dwLowDateTime = blockT<DWORD>::parse(m_Parser);
			dwHighDateTime = blockT<DWORD>::parse(m_Parser);
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.ft = {*dwLowDateTime, *dwHighDateTime}; }
		std::shared_ptr<blockT<DWORD>> dwLowDateTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwHighDateTime = emptyT<DWORD>();
	};

	/* case PT_STRING8 */
	class CountedStringA : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doRuleProcessing)
			{
				str = blockStringA::parse(m_Parser);
				cb->setData(static_cast<DWORD>(str->length()));
			}
			else
			{
				if (m_doNickname)
				{
					m_Parser->advance(sizeof LARGE_INTEGER);
					cb = blockT<DWORD>::parse(m_Parser);
				}
				else
				{
					cb = blockT<DWORD, WORD>::parse(m_Parser);
				}

				str = blockStringA::parse(m_Parser, *cb);
			}
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.lpszA = const_cast<LPSTR>(str->c_str()); }
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
		std::shared_ptr<blockStringA> str = emptySA();
	};

	/* case PT_UNICODE */
	class CountedStringW : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doRuleProcessing)
			{
				str = blockStringW::parse(m_Parser);
				cb->setData(static_cast<DWORD>(str->length()));
			}
			else
			{
				if (m_doNickname)
				{
					m_Parser->advance(sizeof LARGE_INTEGER);
					cb = blockT<DWORD>::parse(m_Parser);
				}
				else
				{
					cb = blockT<DWORD, WORD>::parse(m_Parser);
				}

				str = blockStringW::parse(m_Parser, *cb / sizeof(WCHAR));
			}
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.lpszW = const_cast<LPWSTR>(str->c_str()); }
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
		std::shared_ptr<blockStringW> str = emptySW();
	};

	/* case PT_BINARY */
	class SBinaryBlock : public blockPV
	{
	public:
		void parse(std::shared_ptr<binaryParser>& parser, ULONG ulPropTag)
		{
			blockPV::parse(parser, ulPropTag, false, true);
		}
		operator SBinary() noexcept { return {*cb, const_cast<LPBYTE>(lpb->data())}; }

	private:
		void parse() override
		{
			if (m_doNickname)
			{
				m_Parser->advance(sizeof LARGE_INTEGER);
			}

			if (m_doRuleProcessing || m_doNickname)
			{
				cb = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				cb = blockT<DWORD, WORD>::parse(m_Parser);
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			lpb = blockBytes::parse(m_Parser, *cb);
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.bin = this->operator SBinary(); }
		std::shared_ptr<blockT<ULONG>> cb = emptyT<ULONG>();
		std::shared_ptr<blockBytes> lpb = emptyBB();
	};

	/* case PT_MV_BINARY */
	class SBinaryArrayBlock : public blockPV
	{
	public:
		~SBinaryArrayBlock()
		{
			if (bin) delete[] bin;
		}

	private:
		void parse() override
		{
			if (m_doNickname)
			{
				m_Parser->advance(sizeof LARGE_INTEGER);
				cValues = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				cValues = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (cValues && *cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < *cValues; j++)
				{
					const auto block = std::make_shared<SBinaryBlock>();
					block->parse(m_Parser, m_ulPropTag);
					lpbin.emplace_back(block);
				}
			}

			if (bin) delete[] bin;
			bin = nullptr;
			if (*cValues)
			{
				const auto count = cValues->getData();
				bin = new (std::nothrow) SBinary[count];
				if (bin)
				{
					for (ULONG i = 0; i < count; i++)
					{
						bin[i] = *lpbin[i];
					}
				}
			}
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.MVbin = SBinaryArray{*cValues, bin}; }
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<SBinaryBlock>> lpbin;
		SBinary* bin{};
	};

	/* case PT_MV_STRING8 */
	class StringArrayA : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doNickname)
			{
				m_Parser->advance(sizeof LARGE_INTEGER);
				cValues = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				cValues = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (cValues)
			//if (cValues && cValues < _MaxEntriesLarge)
			{
				lppszA.reserve(*cValues);
				for (ULONG j = 0; j < *cValues; j++)
				{
					lppszA.emplace_back(std::make_shared<blockStringA>(m_Parser));
				}
			}
		}

		// TODO: Populate a SLPSTRArray struct here
		const void getProp(SPropValue& prop) noexcept override { prop.Value.MVszA = {}; }
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringA>> lppszA;
	};

	/* case PT_MV_UNICODE */
	class StringArrayW : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doNickname)
			{
				m_Parser->advance(sizeof LARGE_INTEGER);
				cValues = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				cValues = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (cValues && *cValues < _MaxEntriesLarge)
			{
				lppszW.reserve(*cValues);
				for (ULONG j = 0; j < *cValues; j++)
				{
					lppszW.emplace_back(std::make_shared<blockStringW>(m_Parser));
				}
			}
		}

		// TODO: Populate a SWStringArray struct here
		const void getProp(SPropValue& prop) noexcept override { prop.Value.MVszW = {}; }
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringW>> lppszW;
	};

	/* case PT_I2 */
	class I2BLock : public blockPV
	{
	public:
		std::wstring toNumberAsString() override
		{
			return InterpretNumberAsString(*i, m_ulPropTag, 0, nullptr, nullptr, false);
		}

	private:
		void parse() override
		{
			if (m_doNickname) i = blockT<WORD>::parse(m_Parser); // TODO: This can't be right
			if (m_doNickname) m_Parser->advance(sizeof WORD);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.i = *i; }
		std::shared_ptr<blockT<WORD>> i = emptyT<WORD>();
	};

	/* case PT_LONG */
	class LongBLock : public blockPV
	{
	public:
		std::wstring toNumberAsString() override
		{
			return InterpretNumberAsString(*l, m_ulPropTag, 0, nullptr, nullptr, false);
		}

	private:
		void parse() override
		{
			l = blockT<LONG, DWORD>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.l = *l; }
		std::shared_ptr<blockT<LONG>> l = emptyT<LONG>();
	};

	/* case PT_BOOLEAN */
	class BooleanBlock : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doRuleProcessing)
			{
				b = blockT<WORD, BYTE>::parse(m_Parser);
			}
			else
			{
				b = blockT<WORD>::parse(m_Parser);
			}

			if (m_doNickname) m_Parser->advance(sizeof WORD);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.b = *b; }
		std::shared_ptr<blockT<WORD>> b = emptyT<WORD>();
	};

	/* case PT_R4 */
	class R4BLock : public blockPV
	{
	private:
		void parse() override
		{
			flt = blockT<float>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.flt = *flt; }
		std::shared_ptr<blockT<float>> flt = emptyT<float>();
	};

	/* case PT_DOUBLE */
	class DoubleBlock : public blockPV
	{
	private:
		void parse() override { dbl = blockT<double>::parse(m_Parser); }

		const void getProp(SPropValue& prop) noexcept override { prop.Value.dbl = *dbl; }
		std::shared_ptr<blockT<double>> dbl = emptyT<double>();
	};

	/* case PT_CLSID */
	class CLSIDBlock : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doNickname) m_Parser->advance(sizeof LARGE_INTEGER);
			lpguid = blockT<GUID>::parse(m_Parser);
		}

		const void getProp(SPropValue& prop) noexcept override
		{
			auto guid = lpguid->getData();
			prop.Value.lpguid = &guid;
		}
		std::shared_ptr<blockT<GUID>> lpguid = emptyT<GUID>();
	};

	/* case PT_I8 */
	class I8Block : public blockPV
	{
	public:
		std::wstring toNumberAsString() override
		{
			return InterpretNumberAsString(li->getData().QuadPart, m_ulPropTag, 0, nullptr, nullptr, false);
		}

	private:
		void parse() override { li = blockT<LARGE_INTEGER>::parse(m_Parser); }

		const void getProp(SPropValue& prop) noexcept override { prop.Value.li = li->getData(); }
		std::shared_ptr<blockT<LARGE_INTEGER>> li = emptyT<LARGE_INTEGER>();
	};

	/* case PT_ERROR */
	class ErrorBlock : public blockPV
	{
	private:
		void parse() override
		{
			err = blockT<SCODE, DWORD>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.err = *err; }
		std::shared_ptr<blockT<SCODE>> err = emptyT<SCODE>();
	};

	inline std::shared_ptr<blockPV> getPVParser(ULONG ulPropType)
	{
		switch (ulPropType)
		{
		case PT_I2:
			return std::make_shared<I2BLock>();
		case PT_LONG:
			return std::make_shared<LongBLock>();
		case PT_ERROR:
			return std::make_shared<ErrorBlock>();
		case PT_R4:
			return std::make_shared<R4BLock>();
		case PT_DOUBLE:
			return std::make_shared<DoubleBlock>();
		case PT_BOOLEAN:
			return std::make_shared<BooleanBlock>();
		case PT_I8:
			return std::make_shared<I8Block>();
		case PT_SYSTIME:
			return std::make_shared<FILETIMEBLock>();
		case PT_STRING8:
			return std::make_shared<CountedStringA>();
		case PT_BINARY:
			return std::make_shared<SBinaryBlock>();
		case PT_UNICODE:
			return std::make_shared<CountedStringW>();
		case PT_CLSID:
			return std::make_shared<CLSIDBlock>();
		case PT_MV_STRING8:
			return std::make_shared<StringArrayA>();
		case PT_MV_UNICODE:
			return std::make_shared<StringArrayW>();
		case PT_MV_BINARY:
			return std::make_shared<SBinaryArrayBlock>();
		default:
			return nullptr;
		}
	}
} // namespace smartview