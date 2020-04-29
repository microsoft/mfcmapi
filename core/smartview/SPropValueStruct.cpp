#include <core/stdafx.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
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
			if (m_doNickname) Value.i = blockT<WORD>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof WORD);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
			break;
		case PT_LONG:
			Value.l = blockT<LONG, DWORD>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
			break;
		case PT_ERROR:
			Value.err = blockT<SCODE, DWORD>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
			break;
		case PT_R4:
			Value.flt = blockT<float>::parse(m_Parser);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
			break;
		case PT_DOUBLE:
			Value.dbl = blockT<double>::parse(m_Parser);
			break;
		case PT_BOOLEAN:
			if (m_doRuleProcessing)
			{
				Value.b = blockT<WORD, BYTE>::parse(m_Parser);
			}
			else
			{
				Value.b = blockT<WORD>::parse(m_Parser);
			}

			if (m_doNickname) m_Parser->advance(sizeof WORD);
			if (m_doNickname) m_Parser->advance(sizeof DWORD);
			break;
		case PT_I8:
			Value.li = blockT<LARGE_INTEGER>::parse(m_Parser);
			break;
		case PT_SYSTIME:
			Value.ft.dwHighDateTime = blockT<DWORD>::parse(m_Parser);
			Value.ft.dwLowDateTime = blockT<DWORD>::parse(m_Parser);
			break;
		case PT_STRING8:
			if (m_doRuleProcessing)
			{
				Value.lpszA.str = blockStringA::parse(m_Parser);
				Value.lpszA.cb->setData(static_cast<DWORD>(Value.lpszA.str->length()));
			}
			else
			{
				if (m_doNickname)
				{
					static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
					Value.lpszA.cb = blockT<DWORD>::parse(m_Parser);
				}
				else
				{
					Value.lpszA.cb = blockT<DWORD, WORD>::parse(m_Parser);
				}

				Value.lpszA.str = blockStringA::parse(m_Parser, *Value.lpszA.cb);
			}

			break;
		case PT_BINARY:
			if (m_doNickname)
			{
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
			}

			if (m_doRuleProcessing || m_doNickname)
			{
				Value.bin.cb = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				Value.bin.cb = blockT<DWORD, WORD>::parse(m_Parser);
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			Value.bin.lpb = blockBytes::parse(m_Parser, *Value.bin.cb);
			break;
		case PT_UNICODE:
			if (m_doRuleProcessing)
			{
				Value.lpszW.str = blockStringW::parse(m_Parser);
				Value.lpszW.cb->setData(static_cast<DWORD>(Value.lpszW.str->length()));
			}
			else
			{
				if (m_doNickname)
				{
					static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
					Value.lpszW.cb = blockT<DWORD>::parse(m_Parser);
				}
				else
				{
					Value.lpszW.cb = blockT<DWORD, WORD>::parse(m_Parser);
				}

				Value.lpszW.str = blockStringW::parse(m_Parser, *Value.lpszW.cb / sizeof(WCHAR));
			}
			break;
		case PT_CLSID:
			if (m_doNickname) static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
			Value.lpguid = blockT<GUID>::parse(m_Parser);
			break;
		case PT_MV_STRING8:
			if (m_doNickname)
			{
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
				Value.MVszA.cValues = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				Value.MVszA.cValues = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (Value.MVszA.cValues)
			//if (Value.MVszA.cValues && Value.MVszA.cValues < _MaxEntriesLarge)
			{
				Value.MVszA.lppszA.reserve(*Value.MVszA.cValues);
				for (ULONG j = 0; j < *Value.MVszA.cValues; j++)
				{
					Value.MVszA.lppszA.emplace_back(std::make_shared<blockStringA>(m_Parser));
				}
			}
			break;
		case PT_MV_UNICODE:
			if (m_doNickname)
			{
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
				Value.MVszW.cValues = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				Value.MVszW.cValues = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (Value.MVszW.cValues && *Value.MVszW.cValues < _MaxEntriesLarge)
			{
				Value.MVszW.lppszW.reserve(*Value.MVszW.cValues);
				for (ULONG j = 0; j < *Value.MVszW.cValues; j++)
				{
					Value.MVszW.lppszW.emplace_back(std::make_shared<blockStringW>(m_Parser));
				}
			}
			break;
		case PT_MV_BINARY:
			if (m_doNickname)
			{
				static_cast<void>(m_Parser->advance(sizeof LARGE_INTEGER)); // union
				Value.MVbin.cValues = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				Value.MVbin.cValues = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (Value.MVbin.cValues && *Value.MVbin.cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < *Value.MVbin.cValues; j++)
				{
					Value.MVbin.lpbin.emplace_back(std::make_shared<SBinaryBlock>(m_Parser));
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
			prop.Value.i = *Value.i;
			size = Value.i->getSize();
			offset = Value.i->getOffset();
			break;
		case PT_LONG:
			prop.Value.l = *Value.l;
			size = Value.l->getSize();
			offset = Value.l->getOffset();
			break;
		case PT_R4:
			prop.Value.flt = *Value.flt;
			size = Value.flt->getSize();
			offset = Value.flt->getOffset();
			break;
		case PT_DOUBLE:
			prop.Value.dbl = *Value.dbl;
			size = Value.dbl->getSize();
			offset = Value.dbl->getOffset();
			break;
		case PT_BOOLEAN:
			prop.Value.b = *Value.b;
			size = Value.b->getSize();
			offset = Value.b->getOffset();
			break;
		case PT_I8:
			prop.Value.li = Value.li->getData();
			size = Value.li->getSize();
			offset = Value.li->getOffset();
			break;
		case PT_SYSTIME:
			prop.Value.ft = Value.ft;
			size = Value.ft.getSize();
			offset = Value.ft.getOffset();
			break;
		case PT_STRING8:
			prop.Value.lpszA = const_cast<LPSTR>(Value.lpszA.str->c_str());
			size = Value.lpszA.getSize();
			offset = Value.lpszA.getOffset();
			break;
		case PT_BINARY:
			mapi::setBin(prop) = {*Value.bin.cb, const_cast<LPBYTE>(Value.bin.lpb->data())};
			size = Value.bin.getSize();
			offset = Value.bin.getOffset();
			break;
		case PT_UNICODE:
			prop.Value.lpszW = const_cast<LPWSTR>(Value.lpszW.str->c_str());
			size = Value.lpszW.getSize();
			offset = Value.lpszW.getOffset();
			break;
		case PT_CLSID:
			guid = Value.lpguid->getData();
			prop.Value.lpguid = &guid;
			size = Value.lpguid->getSize();
			offset = Value.lpguid->getOffset();
			break;
		//case PT_MV_STRING8:
		//case PT_MV_UNICODE:
		case PT_MV_BINARY:
			prop.Value.MVbin.cValues = Value.MVbin.cValues->getData();
			prop.Value.MVbin.lpbin = Value.MVbin.getbin();
			break;
		case PT_ERROR:
			prop.Value.err = Value.err->getData();
			size = Value.err->getSize();
			offset = Value.err->getOffset();
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
			return InterpretNumberAsString(*Value.l, *ulPropTag, 0, nullptr, nullptr, false);
		case PT_I2:
			return InterpretNumberAsString(*Value.i, *ulPropTag, 0, nullptr, nullptr, false);
		case PT_I8:
			return InterpretNumberAsString(Value.li->getData().QuadPart, *ulPropTag, 0, nullptr, nullptr, false);
		}

		return strings::emptystring;
	}

} // namespace smartview