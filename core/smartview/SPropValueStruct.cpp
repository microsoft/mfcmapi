#include <core/stdafx.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	/* case PT_SYSTIME */
	class FILETIMEBLock : public blockPV
	{
	private:
		void parse() override
		{
			dwHighDateTime = blockT<DWORD>::parse(m_Parser);
			dwLowDateTime = blockT<DWORD>::parse(m_Parser);
		}

		void getProp(SPropValue& prop) override { prop.Value.ft = {*dwLowDateTime, *dwHighDateTime}; }
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
					static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
					cb = blockT<DWORD>::parse(m_Parser);
				}
				else
				{
					cb = blockT<DWORD, WORD>::parse(m_Parser);
				}

				str = blockStringA::parse(m_Parser, *cb);
			}
		}

		void getProp(SPropValue& prop) override { prop.Value.lpszA = const_cast<LPSTR>(str->c_str()); }
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
					static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
					cb = blockT<DWORD>::parse(m_Parser);
				}
				else
				{
					cb = blockT<DWORD, WORD>::parse(m_Parser);
				}

				str = blockStringW::parse(m_Parser, *cb / sizeof(WCHAR));
			}
		}

		void getProp(SPropValue& prop) override { prop.Value.lpszW = const_cast<LPWSTR>(str->c_str()); }
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
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
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

		void getProp(SPropValue& prop) override { prop.Value.bin = this->operator SBinary(); }
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
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
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
					auto block = std::make_shared<SBinaryBlock>();
					block->parse(m_Parser, m_ulPropTag);
					lpbin.emplace_back(block);
				}
			}
		}

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
	class StringArrayA : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doNickname)
			{
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
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

		void getProp(SPropValue& prop) override { prop.Value.MVszA = SLPSTRArray{}; }
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
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
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

		void getProp(SPropValue& prop) override { prop.Value.MVszW = SWStringArray{}; }
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

		void getProp(SPropValue& prop) override { prop.Value.i = *i; }
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

		void getProp(SPropValue& prop) override { prop.Value.l = *l; }
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

		void getProp(SPropValue& prop) override { prop.Value.b = *b; }
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

		void getProp(SPropValue& prop) override { prop.Value.flt = *flt; }
		std::shared_ptr<blockT<float>> flt = emptyT<float>();
	};

	/* case PT_DOUBLE */
	class DoubleBlock : public blockPV
	{
	private:
		void parse() override { dbl = blockT<double>::parse(m_Parser); }

		void getProp(SPropValue& prop) override { prop.Value.dbl = *dbl; }
		std::shared_ptr<blockT<double>> dbl = emptyT<double>();
	};

	/* case PT_CLSID */
	class CLSIDBlock : public blockPV
	{
	private:
		void parse() override
		{
			if (m_doNickname) static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
			lpguid = blockT<GUID>::parse(m_Parser);
		}

		void getProp(SPropValue& prop) override
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

		void getProp(SPropValue& prop) override { prop.Value.li = li->getData(); }
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
			value = std::make_shared<I2BLock>();
			break;
		case PT_LONG:
			value = std::make_shared<LongBLock>();
			break;
		case PT_ERROR:
			value = std::make_shared<ErrorBlock>();
			break;
		case PT_R4:
			value = std::make_shared<R4BLock>();
			break;
		case PT_DOUBLE:
			value = std::make_shared<DoubleBlock>();
			break;
		case PT_BOOLEAN:
			value = std::make_shared<BooleanBlock>();
			break;
		case PT_I8:
			value = std::make_shared<I8Block>();
			break;
		case PT_SYSTIME:
			value = std::make_shared<FILETIMEBLock>();
			break;
		case PT_STRING8:
			value = std::make_shared<CountedStringA>();
			break;
		case PT_BINARY:
			value = std::make_shared<SBinaryBlock>();
			break;
		case PT_UNICODE:
			value = std::make_shared<CountedStringW>();
			break;
		case PT_CLSID:
			value = std::make_shared<CLSIDBlock>();
			break;
		case PT_MV_STRING8:
			value = std::make_shared<StringArrayA>();
			break;
		case PT_MV_UNICODE:
			value = std::make_shared<StringArrayW>();
			break;
		case PT_MV_BINARY:
			value = std::make_shared<SBinaryArrayBlock>();
			break;
		default:
			break;
		}

		if (value)
		{
			value->parse(m_Parser, *ulPropTag, m_doNickname, m_doRuleProcessing);
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
		if (value)
		{
			auto propString = value->PropBlock();
			if (!propString->empty())
			{
				propRoot->addChild(propString, L"PropString = %1!ws!", propString->c_str());
			}

			auto alt = value->AltPropBlock();
			if (!alt->empty())
			{
				propRoot->addChild(alt, L" AltPropString = %1!ws!", alt->c_str());
			}

			auto szSmartView = value->SmartViewBlock();
			if (!szSmartView->empty())
			{
				propRoot->terminateBlock();
				propRoot->addChild(szSmartView, L"Smart View: %1!ws!", szSmartView->c_str());
			}
		}
	}
} // namespace smartview