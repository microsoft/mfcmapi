#include <core/stdafx.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	SBinaryBlock::SBinaryBlock(const std::shared_ptr<binaryParser>& parser)
	{
		cb = blockT<DWORD>::parse(parser);
		// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
		lpb = blockBytes::parse(parser, *cb);
	}

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
} // namespace smartview