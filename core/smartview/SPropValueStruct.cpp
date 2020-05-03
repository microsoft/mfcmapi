#include <core/stdafx.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	/* case PT_SYSTIME */
	class FILETIMEBLock : public PVBlock
	{
	public:
		FILETIMEBLock(const std::shared_ptr<binaryParser>& parser)
		{
			dwHighDateTime = blockT<DWORD>::parse(parser);
			dwLowDateTime = blockT<DWORD>::parse(parser);
		}
		static std::shared_ptr<FILETIMEBLock>
		parse(const std::shared_ptr<binaryParser>& parser, bool /*doNickname*/, bool /*doRuleProcessing*/)
		{
			return std::make_shared<FILETIMEBLock>(parser);
		}

	private:
		size_t getSize() const noexcept override { return dwLowDateTime->getSize() + dwHighDateTime->getSize(); }
		size_t getOffset() const noexcept override { return dwHighDateTime->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.ft = {*dwLowDateTime, *dwHighDateTime}; }
		std::shared_ptr<blockT<DWORD>> dwLowDateTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwHighDateTime = emptyT<DWORD>();
	};

	/* case PT_STRING8 */
	class CountedStringA : public PVBlock
	{
	public:
		CountedStringA(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
		{
			if (doRuleProcessing)
			{
				str = blockStringA::parse(parser);
				cb->setData(static_cast<DWORD>(str->length()));
			}
			else
			{
				if (doNickname)
				{
					static_cast<void>(parser->advance(sizeof LARGE_INTEGER)); // union
					cb = blockT<DWORD>::parse(parser);
				}
				else
				{
					cb = blockT<DWORD, WORD>::parse(parser);
				}

				str = blockStringA::parse(parser, *cb);
			}
		}
		static std::shared_ptr<CountedStringA>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
		{
			return std::make_shared<CountedStringA>(parser, doNickname, doRuleProcessing);
		}

	private:
		size_t getSize() const noexcept override { return cb->getSize() + str->getSize(); }
		size_t getOffset() const noexcept override { return cb->getOffset() ? cb->getOffset() : str->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.lpszA = const_cast<LPSTR>(str->c_str()); }
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
		std::shared_ptr<blockStringA> str = emptySA();
	};

	/* case PT_UNICODE */
	class CountedStringW : public PVBlock
	{
	public:
		CountedStringW(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
		{
			if (doRuleProcessing)
			{
				str = blockStringW::parse(parser);
				cb->setData(static_cast<DWORD>(str->length()));
			}
			else
			{
				if (doNickname)
				{
					static_cast<void>(parser->advance(sizeof LARGE_INTEGER)); // union
					cb = blockT<DWORD>::parse(parser);
				}
				else
				{
					cb = blockT<DWORD, WORD>::parse(parser);
				}

				str = blockStringW::parse(parser, *cb / sizeof(WCHAR));
			}
		}
		static std::shared_ptr<CountedStringW>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
		{
			return std::make_shared<CountedStringW>(parser, doNickname, doRuleProcessing);
		}

	private:
		size_t getSize() const noexcept override { return cb->getSize() + str->getSize(); }
		size_t getOffset() const noexcept override { return cb->getOffset() ? cb->getOffset() : str->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.lpszW = const_cast<LPWSTR>(str->c_str()); }
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
		std::shared_ptr<blockStringW> str = emptySW();
	};

	/* case PT_BINARY */
	class SBinaryBlock : public PVBlock
	{
	public:
		SBinaryBlock(const std::shared_ptr<binaryParser>& parser) : SBinaryBlock(parser, false, true) {}
		SBinaryBlock(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
		{
			if (doNickname)
			{
				static_cast<void>(parser->advance(sizeof LARGE_INTEGER)); // union
			}

			if (doRuleProcessing || doNickname)
			{
				cb = blockT<DWORD>::parse(parser);
			}
			else
			{
				cb = blockT<DWORD, WORD>::parse(parser);
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			lpb = blockBytes::parse(parser, *cb);
		}
		static std::shared_ptr<SBinaryBlock> parse(const std::shared_ptr<binaryParser>& parser)
		{
			return std::make_shared<SBinaryBlock>(parser);
		}
		static std::shared_ptr<SBinaryBlock>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
		{
			return std::make_shared<SBinaryBlock>(parser, doNickname, doRuleProcessing);
		}

		operator SBinary() noexcept { return {*cb, const_cast<LPBYTE>(lpb->data())}; }

	private:
		size_t getSize() const noexcept { return cb->getSize() + lpb->getSize(); }
		size_t getOffset() const noexcept { return cb->getOffset() ? cb->getOffset() : lpb->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.bin = this->operator SBinary(); }
		std::shared_ptr<blockT<ULONG>> cb = emptyT<ULONG>();
		std::shared_ptr<blockBytes> lpb = emptyBB();
	};

	/* case PT_MV_BINARY */
	class SBinaryArrayBlock : public PVBlock
	{
	public:
		SBinaryArrayBlock(const std::shared_ptr<binaryParser>& parser, bool doNickname)
		{
			if (doNickname)
			{
				static_cast<void>(parser->advance(sizeof LARGE_INTEGER)); // union
				cValues = blockT<DWORD>::parse(parser);
			}
			else
			{
				cValues = blockT<DWORD, WORD>::parse(parser);
			}

			if (cValues && *cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < *cValues; j++)
				{
					lpbin.emplace_back(SBinaryBlock::parse(parser));
				}
			}
		}
		~SBinaryArrayBlock()
		{
			if (bin) delete[] bin;
		}
		static std::shared_ptr<SBinaryArrayBlock>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool /*doRuleProcessing*/)
		{
			return std::make_shared<SBinaryArrayBlock>(parser, doNickname);
		}

	private:
		// TODO: Implement size and offset
		size_t getSize() const noexcept { return {}; }
		size_t getOffset() const noexcept { return {}; }
		void getProp(SPropValue& prop) override
		{
			if (*cValues && !bin)
			{
				auto count = cValues->getData();
				bin = new (std::nothrow) SBinary[count];
				if (bin)
				{
					for (ULONG i = 0; i < count; i++)
					{
						bin[i] = *lpbin[i];
					}
				}
			}

			prop.Value.MVbin = SBinaryArray{*cValues, bin};
		}
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<SBinaryBlock>> lpbin;
		SBinary* bin{};
	};

	/* case PT_MV_STRING8 */
	class StringArrayA : public PVBlock
	{
	public:
		StringArrayA(const std::shared_ptr<binaryParser>& parser, bool doNickname)
		{
			if (doNickname)
			{
				static_cast<void>(parser->advance(sizeof LARGE_INTEGER)); // union
				cValues = blockT<DWORD>::parse(parser);
			}
			else
			{
				cValues = blockT<DWORD, WORD>::parse(parser);
			}

			if (cValues)
			//if (cValues && cValues < _MaxEntriesLarge)
			{
				lppszA.reserve(*cValues);
				for (ULONG j = 0; j < *cValues; j++)
				{
					lppszA.emplace_back(std::make_shared<blockStringA>(parser));
				}
			}
		}
		static std::shared_ptr<StringArrayA>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool /*doRuleProcessing*/)
		{
			return std::make_shared<StringArrayA>(parser, doNickname);
		}

	private:
		// TODO: Implement size and offset
		size_t getSize() const noexcept { return {}; }
		size_t getOffset() const noexcept { return {}; }
		void getProp(SPropValue& prop) override { prop.Value.MVszA = SLPSTRArray{}; }
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringA>> lppszA;
	};

	/* case PT_MV_UNICODE */
	class StringArrayW : public PVBlock
	{
	public:
		StringArrayW(const std::shared_ptr<binaryParser>& parser, bool doNickname)
		{
			if (doNickname)
			{
				static_cast<void>(parser->advance(sizeof LARGE_INTEGER)); // union
				cValues = blockT<DWORD>::parse(parser);
			}
			else
			{
				cValues = blockT<DWORD, WORD>::parse(parser);
			}

			if (cValues && *cValues < _MaxEntriesLarge)
			{
				lppszW.reserve(*cValues);
				for (ULONG j = 0; j < *cValues; j++)
				{
					lppszW.emplace_back(std::make_shared<blockStringW>(parser));
				}
			}
		}
		static std::shared_ptr<StringArrayW>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool /*doRuleProcessing*/)
		{
			return std::make_shared<StringArrayW>(parser, doNickname);
		}

	private:
		// TODO: Implement size and offset
		size_t getSize() const noexcept { return {}; }
		size_t getOffset() const noexcept { return {}; }
		void getProp(SPropValue& prop) override { prop.Value.MVszW = SWStringArray{}; }
		std::shared_ptr<blockT<ULONG>> cValues = emptyT<ULONG>();
		std::vector<std::shared_ptr<blockStringW>> lppszW;
	};

	/* case PT_I2 */
	class I2BLock : public PVBlock
	{
	public:
		I2BLock(const std::shared_ptr<binaryParser>& parser, bool doNickname)
		{
			if (doNickname) i = blockT<WORD>::parse(parser); // TODO: This can't be right
			if (doNickname) parser->advance(sizeof WORD);
			if (doNickname) parser->advance(sizeof DWORD);
		}
		static std::shared_ptr<I2BLock>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool /*doRuleProcessing*/)
		{
			return std::make_shared<I2BLock>(parser, doNickname);
		}

		std::wstring propNum(ULONG ulPropTag) override
		{
			return InterpretNumberAsString(*i, ulPropTag, 0, nullptr, nullptr, false);
		}

	private:
		size_t getSize() const noexcept override { return i->getSize(); }
		size_t getOffset() const noexcept override { return i->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.i = *i; }
		std::shared_ptr<blockT<WORD>> i = emptyT<WORD>();
	};

	/* case PT_LONG */
	class LongBLock : public PVBlock
	{
	public:
		LongBLock(const std::shared_ptr<binaryParser>& parser, bool doNickname)
		{
			l = blockT<LONG, DWORD>::parse(parser);
			if (doNickname) parser->advance(sizeof DWORD);
		}
		static std::shared_ptr<LongBLock>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool /*doRuleProcessing*/)
		{
			return std::make_shared<LongBLock>(parser, doNickname);
		}

		std::wstring propNum(ULONG ulPropTag) override
		{
			return InterpretNumberAsString(*l, ulPropTag, 0, nullptr, nullptr, false);
		}

	private:
		size_t getSize() const noexcept override { return l->getSize(); }
		size_t getOffset() const noexcept override { return l->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.l = *l; }
		std::shared_ptr<blockT<LONG>> l = emptyT<LONG>();
	};

	/* case PT_BOOLEAN */
	class BooleanBLock : public PVBlock
	{
	public:
		BooleanBLock(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
		{
			if (doRuleProcessing)
			{
				b = blockT<WORD, BYTE>::parse(parser);
			}
			else
			{
				b = blockT<WORD>::parse(parser);
			}

			if (doNickname) parser->advance(sizeof WORD);
			if (doNickname) parser->advance(sizeof DWORD);
		}
		static std::shared_ptr<BooleanBLock>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
		{
			return std::make_shared<BooleanBLock>(parser, doNickname, doRuleProcessing);
		}

	private:
		size_t getSize() const noexcept override { return b->getSize(); }
		size_t getOffset() const noexcept override { return b->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.b = *b; }
		std::shared_ptr<blockT<WORD>> b = emptyT<WORD>();
	};

	/* case PT_R4 */
	class R4BLock : public PVBlock
	{
	public:
		R4BLock(const std::shared_ptr<binaryParser>& parser, bool doNickname)
		{
			flt = blockT<float>::parse(parser);
			if (doNickname) parser->advance(sizeof DWORD);
		}
		static std::shared_ptr<R4BLock>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool /*doRuleProcessing*/)
		{
			return std::make_shared<R4BLock>(parser, doNickname);
		}

	private:
		size_t getSize() const noexcept override { return flt->getSize(); }
		size_t getOffset() const noexcept override { return flt->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.flt = *flt; }
		std::shared_ptr<blockT<float>> flt = emptyT<float>();
	};

	/* case PT_DOUBLE */
	class DoubleBlock : public PVBlock
	{
	public:
		DoubleBlock(const std::shared_ptr<binaryParser>& parser) { dbl = blockT<double>::parse(parser); }
		static std::shared_ptr<DoubleBlock>
		parse(const std::shared_ptr<binaryParser>& parser, bool /*doNickname*/, bool /*doRuleProcessing*/)
		{
			return std::make_shared<DoubleBlock>(parser);
		}

	private:
		size_t getSize() const noexcept override { return dbl->getSize(); }
		size_t getOffset() const noexcept override { return dbl->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.dbl = *dbl; }
		std::shared_ptr<blockT<double>> dbl = emptyT<double>();
	};

	/* case PT_CLSID */
	class CLSIDBlock : public PVBlock
	{
	public:
		CLSIDBlock(const std::shared_ptr<binaryParser>& parser, bool doNickname)
		{
			if (doNickname) static_cast<void>(parser->advance(sizeof LARGE_INTEGER)); // union
			lpguid = blockT<GUID>::parse(parser);
		}
		static std::shared_ptr<CLSIDBlock>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool /*doRuleProcessing*/)
		{
			return std::make_shared<CLSIDBlock>(parser, doNickname);
		}

	private:
		size_t getSize() const noexcept override { return lpguid->getSize(); }
		size_t getOffset() const noexcept override { return lpguid->getOffset(); }
		void getProp(SPropValue& prop) override
		{
			GUID guid = lpguid->getData();
			prop.Value.lpguid = &guid;
		}
		std::shared_ptr<blockT<GUID>> lpguid = emptyT<GUID>();
	};

	/* case PT_I8 */
	class I8Block : public PVBlock
	{
	public:
		I8Block(const std::shared_ptr<binaryParser>& parser) { li = blockT<LARGE_INTEGER>::parse(parser); }
		static std::shared_ptr<I8Block>
		parse(const std::shared_ptr<binaryParser>& parser, bool /*doNickname*/, bool /*doRuleProcessing*/)
		{
			return std::make_shared<I8Block>(parser);
		}

		std::wstring propNum(ULONG ulPropTag) override
		{
			return InterpretNumberAsString(li->getData().QuadPart, ulPropTag, 0, nullptr, nullptr, false);
		}

	private:
		size_t getSize() const noexcept override { return li->getSize(); }
		size_t getOffset() const noexcept override { return li->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.li = li->getData(); }
		std::shared_ptr<blockT<LARGE_INTEGER>> li = emptyT<LARGE_INTEGER>();
	};

	/* case PT_ERROR */
	class ErrorBlock : public PVBlock
	{
	public:
		ErrorBlock(const std::shared_ptr<binaryParser>& parser, bool doNickname)
		{
			err = blockT<SCODE, DWORD>::parse(parser);
			if (doNickname) parser->advance(sizeof DWORD);
		}
		static std::shared_ptr<ErrorBlock>
		parse(const std::shared_ptr<binaryParser>& parser, bool doNickname, bool /*doRuleProcessing*/)
		{
			return std::make_shared<ErrorBlock>(parser, doNickname);
		}

	private:
		size_t getSize() const noexcept override { return err->getSize(); }
		size_t getOffset() const noexcept override { return err->getOffset(); }
		void getProp(SPropValue& prop) override { prop.Value.err = *err; }
		std::shared_ptr<blockT<SCODE>> err = emptyT<SCODE>();
	};

	void SPropValueStruct::parse()
	{
		const auto ulCurrOffset = m_Parser->getOffset();

		PropType = blockT<WORD>::parse(m_Parser);
		PropID = blockT<WORD>::parse(m_Parser);

		ulPropTag = blockT<ULONG>::create(
			PROP_TAG(*PropType, *PropID), PropType->getSize() + PropID->getSize(), PropType->getOffset());

		if (m_doNickname) static_cast<void>(m_Parser->advance(sizeof DWORD)); // reserved

		switch (*PropType)
		{
		case PT_I2:
			value = I2BLock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_LONG:
			value = LongBLock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_ERROR:
			value = ErrorBlock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_R4:
			value = R4BLock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_DOUBLE:
			value = DoubleBlock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_BOOLEAN:
			value = BooleanBLock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_I8:
			value = I8Block::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_SYSTIME:
			value = FILETIMEBLock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_STRING8:
			value = CountedStringA::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_BINARY:
			value = SBinaryBlock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_UNICODE:
			value = CountedStringW::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_CLSID:
			value = CLSIDBlock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_MV_STRING8:
			value = StringArrayA::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_MV_UNICODE:
			value = StringArrayW::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		case PT_MV_BINARY:
			value = SBinaryArrayBlock::parse(m_Parser, m_doNickname, m_doRuleProcessing);
			break;
		default:
			break;
		}
	}

	void SPropValueStruct::parseBlocks()
	{
		auto propRoot = std::make_shared<block>();
		addChild(propRoot);
		propRoot->setText(L"Property[%1!d!]\r\n", m_index);
		propRoot->addChild(ulPropTag, L"Property = 0x%1!08X!", ulPropTag->getData());

		auto propTagNames = proptags::PropTagToPropName(*ulPropTag, false);
		if (!propTagNames.bestGuess.empty())
		{
			// TODO: Add this as a child of ulPropTag
			propRoot->terminateBlock();
			propRoot->addHeader(L"Name: %1!ws!", propTagNames.bestGuess.c_str());
		}

		if (!propTagNames.otherMatches.empty())
		{
			// TODO: Add this as a child of ulPropTag
			propRoot->terminateBlock();
			propRoot->addHeader(L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
		}

		propRoot->terminateBlock();
		auto propString = PropBlock();
		if (!propString->empty())
		{
			propRoot->addChild(propString, L"PropString = %1!ws!", propString->c_str());
		}

		auto alt = AltPropBlock();
		if (!alt->empty())
		{
			propRoot->addChild(alt, L" AltPropString = %1!ws!", alt->c_str());
		}

		auto szSmartView = SmartViewBlock();
		if (!szSmartView->empty())
		{
			propRoot->terminateBlock();
			propRoot->addChild(szSmartView, L"Smart View: %1!ws!", szSmartView->c_str());
		}
	}

	_Check_return_ std::wstring SPropValueStruct::PropNum() const
	{
		return value ? value->propNum(*ulPropTag) : strings::emptystring;
	}
} // namespace smartview