#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/SmartView.h>
#include <core/property/parseProperty.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>
#include <core/addin/mfcmapi.h>
#include <core/addin/addin.h>

namespace smartview
{
	class blockPV : public block
	{
	public:
		void init(ULONG ulPropTag, bool doNickname, bool doRuleProcessing, bool bMVRow)
		{
			m_doNickname = doNickname;
			m_doRuleProcessing = doRuleProcessing;
			m_ulPropTag = ulPropTag;
			svParser = FindSmartViewParserForProp(ulPropTag, nullptr, nullptr, nullptr, false, bMVRow);
		}

	protected:
		bool m_doNickname{};
		bool m_doRuleProcessing{};
		ULONG m_ulPropTag{};
		parserType svParser{parserType::NOPARSING};

	private:
		void parse() override = 0;
		void parseBlocks() override
		{
			const auto size = parser->getOffset() - getOffset();
			auto prop = SPropValue{m_ulPropTag, 0, {}};
			getProp(prop);

			auto propString = std::wstring{};
			auto altPropString = std::wstring{};
			property::parseProperty(&prop, &propString, &altPropString);

			const auto propBlock =
				blockStringW::parse(strings::RemoveInvalidCharactersW(propString, false), size, getOffset());
			if (!propBlock->empty())
			{
				addChild(propBlock, L"PropString = %1!ws!", propBlock->c_str());
			}

			const auto altPropBlock =
				blockStringW::parse(strings::RemoveInvalidCharactersW(altPropString, false), size, getOffset());
			if (!altPropBlock->empty())
			{
				addChild(altPropBlock, L"AltPropString = %1!ws!", altPropBlock->c_str());
			}

			const auto smartView = toSmartView();
			if (smartView->hasData())
			{
				addLabeledChild(L"Smart View", smartView);
			}
		}
		virtual const void getProp(SPropValue& prop) noexcept = 0;
		virtual std::shared_ptr<block> toSmartView() { return emptySW(); }
	};

	/* case PT_SYSTIME */
	class FILETIMEBLock : public blockPV
	{
	private:
		void parse() override
		{
			dwLowDateTime = blockT<DWORD>::parse(parser);
			dwHighDateTime = blockT<DWORD>::parse(parser);
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
				str = blockStringA::parse(parser);
				cb->setData(static_cast<DWORD>(str->length()));
			}
			else
			{
				if (m_doNickname)
				{
					parser->advance(sizeof LARGE_INTEGER);
					cb = blockT<DWORD>::parse(parser);
				}
				else
				{
					cb = blockT<DWORD>::parse<WORD>(parser);
				}

				str = blockStringA::parse(parser, *cb);
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
				str = blockStringW::parse(parser);
				cb->setData(static_cast<DWORD>(str->length()));
			}
			else
			{
				if (m_doNickname)
				{
					parser->advance(sizeof LARGE_INTEGER);
					cb = blockT<DWORD>::parse(parser);
				}
				else
				{
					cb = blockT<DWORD>::parse<WORD>(parser);
				}

				str = blockStringW::parse(parser, *cb / sizeof(WCHAR));
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
		operator SBinary() noexcept { return {*cb, const_cast<LPBYTE>(lpb->data())}; }

		std::shared_ptr<block> toSmartView() override
		{
			auto svp = InterpretBinary(this->operator SBinary(), svParser, nullptr);
			svp->shiftOffset(lpb->getOffset());
			return svp;
		}

	private:
		void parse() override
		{
			if (m_doNickname)
			{
				parser->advance(sizeof LARGE_INTEGER);
			}

			if (m_doRuleProcessing || m_doNickname)
			{
				cb = blockT<DWORD>::parse(parser);
			}
			else
			{
				cb = blockT<DWORD>::parse<WORD>(parser);
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			lpb = blockBytes::parse(parser, *cb);
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
				parser->advance(sizeof LARGE_INTEGER);
				cValues = blockT<DWORD>::parse(parser);
			}
			else
			{
				cValues = blockT<DWORD>::parse<WORD>(parser);
			}

			if (cValues && *cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < *cValues; j++)
				{
					const auto block = std::make_shared<SBinaryBlock>();
					block->init(CHANGE_PROP_TYPE(m_ulPropTag, PT_BINARY), false, true, true);
					block->block::parse(parser, false);
					lpbin.emplace_back(block);
				}
			}

			if (bin) delete[] bin;
			bin = nullptr;
			if (lpbin.size())
			{
				const auto count = lpbin.size();
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

		std::shared_ptr<block> toSmartView() override
		{
			auto smartview = block::create();
			auto i = 0;
			for (const auto& b : lpbin)
			{
				smartview->addLabeledChild(strings::format(L"Row %d", i++), b->toSmartView());
			}

			return smartview;
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.MVbin = SBinaryArray{*cValues, bin}; }
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<SBinaryBlock>> lpbin;
		SBinary* bin{};
	};

	/* case PT_MV_STRING8 */
	class StringArrayA : public blockPV
	{
	public:
		~StringArrayA()
		{
			if (strings) delete[] strings;
		}

	private:
		void parse() override
		{
			if (m_doNickname)
			{
				parser->advance(sizeof LARGE_INTEGER);
				cValues = blockT<DWORD>::parse(parser);
			}
			else
			{
				cValues = blockT<DWORD>::parse(parser);
			}

			if (cValues && *cValues < _MaxEntriesLarge)
			{
				lppszA.reserve(*cValues);
				for (ULONG j = 0; j < *cValues; j++)
				{
					lppszA.emplace_back(blockStringA::parse(parser));
				}
			}

			if (strings) delete[] strings;
			strings = nullptr;
			if (lppszA.size())
			{
				const auto count = lppszA.size();
				strings = new (std::nothrow) LPSTR[count];
				if (strings)
				{
					for (ULONG i = 0; i < count; i++)
					{
						strings[i] = const_cast<LPSTR>(lppszA[i]->c_str());
					}
				}
			}
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.MVszA = {}; }
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringA>> lppszA;
		LPSTR* strings{};
	};

	/* case PT_MV_UNICODE */
	class StringArrayW : public blockPV
	{
	public:
		~StringArrayW()
		{
			if (strings) delete[] strings;
		}

	private:
		void parse() override
		{
			if (m_doNickname)
			{
				parser->advance(sizeof LARGE_INTEGER);
				cValues = blockT<DWORD>::parse(parser);
			}
			else
			{
				cValues = blockT<DWORD>::parse(parser);
			}

			if (cValues && *cValues < _MaxEntriesLarge)
			{
				lppszW.reserve(*cValues);
				for (ULONG j = 0; j < *cValues; j++)
				{
					lppszW.emplace_back(blockStringW::parse(parser));
				}
			}

			if (strings) delete[] strings;
			strings = nullptr;
			if (lppszW.size())
			{
				const auto count = lppszW.size();
				strings = new (std::nothrow) LPWSTR[count];
				if (strings)
				{
					for (ULONG i = 0; i < count; i++)
					{
						strings[i] = const_cast<LPWSTR>(lppszW[i]->c_str());
					}
				}
			}
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.MVszW = {*cValues, strings}; }
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringW>> lppszW;
		LPWSTR* strings{};
	};

	/* case PT_I2 */
	class I2BLock : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doNickname) i = blockT<WORD>::parse(parser); // TODO: This can't be right
			if (m_doNickname) parser->advance(sizeof WORD);
			if (m_doNickname) parser->advance(sizeof DWORD);
		}

		std::shared_ptr<block> toSmartView() override
		{
			return blockStringW::create(InterpretNumberAsString(*i, m_ulPropTag, 0, nullptr, nullptr, true));
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.i = *i; }
		std::shared_ptr<blockT<WORD>> i = emptyT<WORD>();
	};

	/* case PT_LONG */
	class LongBLock : public blockPV
	{
	private:
		void parse() override
		{
			l = blockT<LONG>::parse<DWORD>(parser);
			if (m_doNickname) parser->advance(sizeof DWORD);
		}

		std::shared_ptr<block> toSmartView() override
		{
			return blockStringW::create(InterpretNumberAsString(*l, m_ulPropTag, 0, nullptr, nullptr, true));
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
				b = blockT<WORD>::parse<BYTE>(parser);
			}
			else
			{
				b = blockT<WORD>::parse(parser);
			}

			if (m_doNickname) parser->advance(sizeof WORD);
			if (m_doNickname) parser->advance(sizeof DWORD);
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
			flt = blockT<float>::parse(parser);
			if (m_doNickname) parser->advance(sizeof DWORD);
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.flt = *flt; }
		std::shared_ptr<blockT<float>> flt = emptyT<float>();
	};

	/* case PT_DOUBLE */
	class DoubleBlock : public blockPV
	{
	private:
		void parse() override { dbl = blockT<double>::parse(parser); }

		const void getProp(SPropValue& prop) noexcept override { prop.Value.dbl = *dbl; }
		std::shared_ptr<blockT<double>> dbl = emptyT<double>();
	};

	/* case PT_CLSID */
	class CLSIDBlock : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doNickname) parser->advance(sizeof LARGE_INTEGER);
			lpguid = blockT<GUID>::parse(parser);
		}

		const void getProp(SPropValue& prop) noexcept override
		{
			// If we use getData, we create a stack copy of the GUID, which goes away.
			// So we need to use getDataAddress to get a pointer to the GUID.
			auto guid = lpguid->getDataAddress();
			prop.Value.lpguid = const_cast<LPGUID>(guid);
		}
		std::shared_ptr<blockT<GUID>> lpguid = emptyT<GUID>();
	};

	/* case PT_I8 */
	class I8Block : public blockPV
	{
	private:
		void parse() override { li = blockT<LARGE_INTEGER>::parse(parser); }

		std::shared_ptr<block> toSmartView() override
		{
			return blockStringW::create(
				InterpretNumberAsString(li->getData().QuadPart, m_ulPropTag, 0, nullptr, nullptr, true));
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.li = li->getData(); }
		std::shared_ptr<blockT<LARGE_INTEGER>> li = emptyT<LARGE_INTEGER>();
	};

	/* case PT_ERROR */
	class ErrorBlock : public blockPV
	{
	private:
		void parse() override
		{
			err = blockT<SCODE>::parse<DWORD>(parser);
			if (m_doNickname) parser->advance(sizeof DWORD);
		}

		const void getProp(SPropValue& prop) noexcept override { prop.Value.err = *err; }
		std::shared_ptr<blockT<SCODE>> err = emptyT<SCODE>();
	};

	inline std::shared_ptr<blockPV> getPVParser(ULONG ulPropTag, bool doNickname, bool doRuleProcessing)
	{
		auto ret = std::shared_ptr<blockPV>{};
		switch (PROP_TYPE(ulPropTag))
		{
		case PT_I2:
			ret = std::make_shared<I2BLock>();
			break;
		case PT_LONG:
			ret = std::make_shared<LongBLock>();
			break;
		case PT_ERROR:
			ret = std::make_shared<ErrorBlock>();
			break;
		case PT_R4:
			ret = std::make_shared<R4BLock>();
			break;
		case PT_DOUBLE:
			ret = std::make_shared<DoubleBlock>();
			break;
		case PT_BOOLEAN:
			ret = std::make_shared<BooleanBlock>();
			break;
		case PT_I8:
			ret = std::make_shared<I8Block>();
			break;
		case PT_SYSTIME:
			ret = std::make_shared<FILETIMEBLock>();
			break;
		case PT_STRING8:
			ret = std::make_shared<CountedStringA>();
			break;
		case PT_BINARY:
			ret = std::make_shared<SBinaryBlock>();
			break;
		case PT_UNICODE:
			ret = std::make_shared<CountedStringW>();
			break;
		case PT_CLSID:
			ret = std::make_shared<CLSIDBlock>();
			break;
		case PT_MV_STRING8:
			ret = std::make_shared<StringArrayA>();
			break;
		case PT_MV_UNICODE:
			ret = std::make_shared<StringArrayW>();
			break;
		case PT_MV_BINARY:
			ret = std::make_shared<SBinaryArrayBlock>();
			break;
		default:
			return nullptr;
		}

		ret->init(ulPropTag, doNickname, doRuleProcessing, false);
		return ret;
	}
} // namespace smartview