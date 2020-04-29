#include <core/stdafx.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	class FILETIMEBLock
	{
	public:
		FILETIMEBLock(const std::shared_ptr<binaryParser>& parser)
		{
			dwHighDateTime = blockT<DWORD>::parse(parser);
			dwLowDateTime = blockT<DWORD>::parse(parser);
		}
		static std::shared_ptr<FILETIMEBLock> parse(const std::shared_ptr<binaryParser>& parser)
		{
			return std::make_shared<FILETIMEBLock>(parser);
		}

		operator FILETIME() const noexcept { return FILETIME{*dwLowDateTime, *dwHighDateTime}; }
		size_t getSize() const noexcept { return dwLowDateTime->getSize() + dwHighDateTime->getSize(); }
		size_t getOffset() const noexcept { return dwHighDateTime->getOffset(); }

	private:
		std::shared_ptr<blockT<DWORD>> dwLowDateTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwHighDateTime = emptyT<DWORD>();
	};

	class CountedStringA
	{
	public:
		CountedStringA(const std::shared_ptr<binaryParser>& parser, bool doRuleProcessing, bool doNickname)
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
		parse(const std::shared_ptr<binaryParser>& parser, bool doRuleProcessing, bool doNickname)
		{
			return std::make_shared<CountedStringA>(parser, doRuleProcessing, doNickname);
		}

		operator LPSTR() const noexcept { return const_cast<LPSTR>(str->c_str()); }
		size_t getSize() const noexcept { return cb->getSize() + str->getSize(); }
		size_t getOffset() const noexcept { return cb->getOffset() ? cb->getOffset() : str->getOffset(); }

	private:
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
		std::shared_ptr<blockStringA> str = emptySA();
	};

	void SPropValueStruct::parse()
	{
		const auto ulCurrOffset = m_Parser->getOffset();

		PropType = blockT<WORD>::parse(m_Parser);
		PropID = blockT<WORD>::parse(m_Parser);

		ulPropTag = blockT<ULONG>::create(
			PROP_TAG(*PropType, *PropID), PropType->getSize() + PropID->getSize(), PropType->getOffset());
		dwAlignPad = 0;

		if (m_doNickname) static_cast<void>(m_Parser->advance(sizeof DWORD)); // reserved

		switch (*PropType)
		{
		case PT_I2:
			// TODO: Insert proper property struct parsing here
			if (m_doNickname) i = blockT<WORD>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof WORD);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
			break;
		case PT_LONG:
			l = blockT<LONG, DWORD>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
			break;
		case PT_ERROR:
			err = blockT<SCODE, DWORD>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
			break;
		case PT_R4:
			flt = blockT<float>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
			break;
		case PT_DOUBLE:
			dbl = blockT<double>::parse(m_Parser);
			break;
		case PT_BOOLEAN:
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
			break;
		case PT_I8:
			li = blockT<LARGE_INTEGER>::parse(m_Parser);
			break;
		case PT_SYSTIME:
			ft = FILETIMEBLock::parse(m_Parser);
			break;
		case PT_STRING8:
			lpszA = CountedStringA::parse(m_Parser, m_doRuleProcessing, m_doNickname);
			break;
		case PT_BINARY:
			if (m_doNickname)
			{
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
			}

			if (m_doRuleProcessing || m_doNickname)
			{
				bin.cb = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				bin.cb = blockT<DWORD, WORD>::parse(m_Parser);
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			bin.lpb = blockBytes::parse(m_Parser, *bin.cb);
			break;
		case PT_UNICODE:
			if (m_doRuleProcessing)
			{
				lpszW.str = blockStringW::parse(m_Parser);
				lpszW.cb->setData(static_cast<DWORD>(lpszW.str->length()));
			}
			else
			{
				if (m_doNickname)
				{
					static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
					lpszW.cb = blockT<DWORD>::parse(m_Parser);
				}
				else
				{
					lpszW.cb = blockT<DWORD, WORD>::parse(m_Parser);
				}

				lpszW.str = blockStringW::parse(m_Parser, *lpszW.cb / sizeof(WCHAR));
			}
			break;
		case PT_CLSID:
			if (m_doNickname) static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
			lpguid = blockT<GUID>::parse(m_Parser);
			break;
		case PT_MV_STRING8:
			if (m_doNickname)
			{
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
				MVszA.cValues = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				MVszA.cValues = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (MVszA.cValues)
			//if (MVszA.cValues && MVszA.cValues < _MaxEntriesLarge)
			{
				MVszA.lppszA.reserve(*MVszA.cValues);
				for (ULONG j = 0; j < *MVszA.cValues; j++)
				{
					MVszA.lppszA.emplace_back(std::make_shared<blockStringA>(m_Parser));
				}
			}
			break;
		case PT_MV_UNICODE:
			if (m_doNickname)
			{
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
				MVszW.cValues = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				MVszW.cValues = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (MVszW.cValues && *MVszW.cValues < _MaxEntriesLarge)
			{
				MVszW.lppszW.reserve(*MVszW.cValues);
				for (ULONG j = 0; j < *MVszW.cValues; j++)
				{
					MVszW.lppszW.emplace_back(std::make_shared<blockStringW>(m_Parser));
				}
			}
			break;
		case PT_MV_BINARY:
			if (m_doNickname)
			{
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
				MVbin.cValues = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				MVbin.cValues = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (MVbin.cValues && *MVbin.cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < *MVbin.cValues; j++)
				{
					MVbin.lpbin.emplace_back(std::make_shared<SBinaryBlock>(m_Parser));
				}
			}
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

	void SPropValueStruct::EnsurePropBlocks()
	{
		if (propStringsGenerated) return;
		auto prop = SPropValue{};
		auto size = size_t{};
		auto offset = size_t{};
		prop.ulPropTag = *ulPropTag;
		prop.dwAlignPad = dwAlignPad;
		switch (*PropType)
		{
		case PT_I2:
			prop.Value.i = *i;
			size = i->getSize();
			offset = i->getOffset();
			break;
		case PT_LONG:
			prop.Value.l = *l;
			size = l->getSize();
			offset = l->getOffset();
			break;
		case PT_R4:
			prop.Value.flt = *flt;
			size = flt->getSize();
			offset = flt->getOffset();
			break;
		case PT_DOUBLE:
			prop.Value.dbl = *dbl;
			size = dbl->getSize();
			offset = dbl->getOffset();
			break;
		case PT_BOOLEAN:
			prop.Value.b = *b;
			size = b->getSize();
			offset = b->getOffset();
			break;
		case PT_I8:
			prop.Value.li = li->getData();
			size = li->getSize();
			offset = li->getOffset();
			break;
		case PT_SYSTIME:
			if (ft)
			{
				prop.Value.ft = *ft;
				size = ft->getSize();
				offset = ft->getOffset();
			}
			break;
		case PT_STRING8:
			if (lpszA)
			{
				prop.Value.lpszA = *lpszA;
				size = lpszA->getSize();
				offset = lpszA->getOffset();
			}
			break;
		case PT_BINARY:
			mapi::setBin(prop) = {*bin.cb, const_cast<LPBYTE>(bin.lpb->data())};
			size = bin.getSize();
			offset = bin.getOffset();
			break;
		case PT_UNICODE:
			prop.Value.lpszW = const_cast<LPWSTR>(lpszW.str->c_str());
			size = lpszW.getSize();
			offset = lpszW.getOffset();
			break;
		case PT_CLSID:
			guid = lpguid->getData();
			prop.Value.lpguid = &guid;
			size = lpguid->getSize();
			offset = lpguid->getOffset();
			break;
		//case PT_MV_STRING8:
		//case PT_MV_UNICODE:
		case PT_MV_BINARY:
			prop.Value.MVbin.cValues = MVbin.cValues->getData();
			prop.Value.MVbin.lpbin = MVbin.getbin();
			break;
		case PT_ERROR:
			prop.Value.err = err->getData();
			size = err->getSize();
			offset = err->getOffset();
			break;
		}

		auto propString = std::wstring{};
		auto altPropString = std::wstring{};
		property::parseProperty(&prop, &propString, &altPropString);

		propBlock = std::make_shared<blockStringW>(strings::RemoveInvalidCharactersW(propString, false), size, offset);

		altPropBlock =
			std::make_shared<blockStringW>(strings::RemoveInvalidCharactersW(altPropString, false), size, offset);

		const auto smartViewString = parsePropertySmartView(&prop, nullptr, nullptr, nullptr, false, false);
		smartViewBlock = std::make_shared<blockStringW>(smartViewString, size, offset);

		propStringsGenerated = true;
	}

	_Check_return_ std::wstring SPropValueStruct::PropNum() const
	{
		switch (PROP_TYPE(*ulPropTag))
		{
		case PT_LONG:
			return InterpretNumberAsString(*l, *ulPropTag, 0, nullptr, nullptr, false);
		case PT_I2:
			return InterpretNumberAsString(*i, *ulPropTag, 0, nullptr, nullptr, false);
		case PT_I8:
			return InterpretNumberAsString(li->getData().QuadPart, *ulPropTag, 0, nullptr, nullptr, false);
		}

		return strings::emptystring;
	}

} // namespace smartview