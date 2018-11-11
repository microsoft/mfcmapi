#include <StdAfx.h>
#include <Interpret/SmartView/PropertiesStruct.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/SmartView/SmartView.h>

namespace smartview
{
	void PropertiesStruct::Parse()
	{
		// For consistancy with previous parsings, we'll refuse to parse if asked to parse more than _MaxEntriesSmall
		// However, we may want to reconsider this choice.
		if (m_MaxEntries > _MaxEntriesSmall) return;

		DWORD dwPropCount = 0;

		// If we have a non-default max, it was computed elsewhere and we do expect to have that many entries. So we can reserve.
		if (m_MaxEntries != _MaxEntriesSmall)
		{
			m_Props.reserve(m_MaxEntries);
		}

		for (;;)
		{
			if (dwPropCount >= m_MaxEntries) break;
			m_Props.push_back(BinToSPropValueStruct());
			if (!m_Parser.RemainingBytes()) break;
			dwPropCount++;
		}
	}

	void PropertiesStruct::ParseBlocks()
	{
		auto i = 0;
		for (auto prop : m_Props)
		{
			if (i != 0)
			{
				addBlankLine();
			}

			addHeader(L"Property[%1!d!]\r\n", i++);
			addBlock(prop.ulPropTag, L"Property = 0x%1!08X!", prop.ulPropTag.getData());

			auto propTagNames = interpretprop::PropTagToPropName(prop.ulPropTag, false);
			if (!propTagNames.bestGuess.empty())
			{
				terminateBlock();
				addBlock(prop.ulPropTag, L"Name: %1!ws!", propTagNames.bestGuess.c_str());
			}

			if (!propTagNames.otherMatches.empty())
			{
				terminateBlock();
				addBlock(prop.ulPropTag, L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
			}

			// TODO: get proper blocks here
			terminateBlock();
			addHeader(L"PropString = %1!ws! ", prop.PropString().c_str());
			addHeader(L"AltPropString = %1!ws!", prop.AltPropString().c_str());

			auto sProp = prop.getData();
			auto szSmartView = InterpretPropSmartView(&sProp, nullptr, nullptr, nullptr, false, false);

			if (!szSmartView.empty())
			{
				// TODO: proper blocks here
				terminateBlock();
				addHeader(L"Smart View: %1!ws!", szSmartView.c_str());
			}
		}
	}

	_Check_return_ SPropValueStruct PropertiesStruct::BinToSPropValueStruct()
	{
		const auto ulCurrOffset = m_Parser.GetCurrentOffset();

		auto prop = SPropValueStruct{};
		prop.PropType = m_Parser.Get<WORD>();
		prop.PropID = m_Parser.Get<WORD>();

		prop.ulPropTag = PROP_TAG(prop.PropType, prop.PropID);
		prop.ulPropTag.setSize(prop.PropType.getSize() + prop.PropID.getSize());
		prop.ulPropTag.setOffset(prop.PropType.getOffset());
		prop.dwAlignPad = 0;

		if (m_NickName) (void) m_Parser.Get<DWORD>(); // reserved

		switch (prop.PropType)
		{
		case PT_I2:
			// TODO: Insert proper property struct parsing here
			if (m_NickName) prop.Value.i = m_Parser.Get<WORD>();
			if (m_NickName) m_Parser.Get<WORD>();
			if (m_NickName) m_Parser.Get<DWORD>();
			break;
		case PT_LONG:
			prop.Value.l = m_Parser.Get<DWORD>();
			if (m_NickName) m_Parser.Get<DWORD>();
			break;
		case PT_ERROR:
			prop.Value.err = m_Parser.Get<DWORD>();
			if (m_NickName) m_Parser.Get<DWORD>();
			break;
		case PT_R4:
			prop.Value.flt = m_Parser.Get<float>();
			if (m_NickName) m_Parser.Get<DWORD>();
			break;
		case PT_DOUBLE:
			prop.Value.dbl = m_Parser.Get<double>();
			break;
		case PT_BOOLEAN:
			if (m_RuleCondition)
			{
				prop.Value.b = m_Parser.Get<BYTE>();
			}
			else
			{
				prop.Value.b = m_Parser.Get<WORD>();
			}

			if (m_NickName) m_Parser.Get<WORD>();
			if (m_NickName) m_Parser.Get<DWORD>();
			break;
		case PT_I8:
			prop.Value.li = m_Parser.Get<LARGE_INTEGER>();
			break;
		case PT_SYSTIME:
			prop.Value.ft.dwHighDateTime = m_Parser.Get<DWORD>();
			prop.Value.ft.dwLowDateTime = m_Parser.Get<DWORD>();
			break;
		case PT_STRING8:
			if (m_RuleCondition)
			{
				prop.Value.lpszA.str = m_Parser.GetStringA();
				prop.Value.lpszA.cb = static_cast<ULONG>(prop.Value.lpszA.str.length());
			}
			else
			{
				if (m_NickName)
				{
					(void) m_Parser.Get<LARGE_INTEGER>(); // union
					prop.Value.lpszA.cb = m_Parser.Get<DWORD>();
				}
				else
				{
					prop.Value.lpszA.cb = m_Parser.Get<WORD>();
				}

				prop.Value.lpszA.str = m_Parser.GetStringA(prop.Value.lpszA.cb);
			}

			break;
		case PT_BINARY:
			if (m_NickName)
			{
				(void) m_Parser.Get<LARGE_INTEGER>(); // union
			}

			if (m_RuleCondition || m_NickName)
			{
				prop.Value.bin.cb = m_Parser.Get<DWORD>();
			}
			else
			{
				prop.Value.bin.cb = m_Parser.Get<WORD>();
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			prop.Value.bin.lpb = m_Parser.GetBYTES(prop.Value.bin.cb);
			break;
		case PT_UNICODE:
			if (m_RuleCondition)
			{
				prop.Value.lpszW.str = m_Parser.GetStringW();
				prop.Value.lpszW.cb = static_cast<ULONG>(prop.Value.lpszW.str.length());
			}
			else
			{
				if (m_NickName)
				{
					(void) m_Parser.Get<LARGE_INTEGER>(); // union
					prop.Value.lpszW.cb = m_Parser.Get<DWORD>();
				}
				else
				{
					prop.Value.lpszW.cb = m_Parser.Get<WORD>();
				}

				prop.Value.lpszW.str = m_Parser.GetStringW(prop.Value.lpszW.cb / sizeof(WCHAR));
			}
			break;
		case PT_CLSID:
			if (m_NickName) (void) m_Parser.Get<LARGE_INTEGER>(); // union
			prop.Value.lpguid = m_Parser.Get<GUID>();
			break;
		case PT_MV_STRING8:
			if (m_NickName)
			{
				(void) m_Parser.Get<LARGE_INTEGER>(); // union
				prop.Value.MVszA.cValues = m_Parser.Get<DWORD>();
			}
			else
			{
				prop.Value.MVszA.cValues = m_Parser.Get<WORD>();
			}

			if (prop.Value.MVszA.cValues)
			//if (prop.Value.MVszA.cValues && prop.Value.MVszA.cValues < _MaxEntriesLarge)
			{
				prop.Value.MVszA.lppszA.reserve(prop.Value.MVszA.cValues);
				for (ULONG j = 0; j < prop.Value.MVszA.cValues; j++)
				{
					prop.Value.MVszA.lppszA.push_back(m_Parser.GetStringA());
				}
			}
			break;
		case PT_MV_UNICODE:
			if (m_NickName)
			{
				(void) m_Parser.Get<LARGE_INTEGER>(); // union
				prop.Value.MVszW.cValues = m_Parser.Get<DWORD>();
			}
			else
			{
				prop.Value.MVszW.cValues = m_Parser.Get<WORD>();
			}

			if (prop.Value.MVszW.cValues)
			//if (prop.Value.MVszW.cValues && prop.Value.MVszW.cValues < _MaxEntriesLarge)
			{
				prop.Value.MVszW.lppszW.reserve(prop.Value.MVszW.cValues);
				for (ULONG j = 0; j < prop.Value.MVszW.cValues; j++)
				{
					prop.Value.MVszW.lppszW.emplace_back(m_Parser.GetStringW());
				}
			}
			break;
		case PT_MV_BINARY:
			if (m_NickName)
			{
				(void) m_Parser.Get<LARGE_INTEGER>(); // union
				prop.Value.MVbin.cValues = m_Parser.Get<DWORD>();
			}
			else
			{
				prop.Value.MVbin.cValues = m_Parser.Get<WORD>();
			}

			if (prop.Value.MVbin.cValues && prop.Value.MVbin.cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < prop.Value.MVbin.cValues; j++)
				{
					auto bin = SBinaryBlock{};
					bin.cb = m_Parser.Get<DWORD>();
					// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
					bin.lpb = m_Parser.GetBYTES(bin.cb);
					prop.Value.MVbin.lpbin.push_back(bin);
				}
			}
			break;
		default:
			break;
		}

		if (ulCurrOffset == m_Parser.GetCurrentOffset())
		{
			// We didn't advance at all - we should return nothing.
			return {};
		}

		return prop;
	}
} // namespace smartview