#include <StdAfx.h>
#include <Interpret/SmartView/PropertyStruct.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/SmartView/SmartView.h>

namespace smartview
{
	PropertyStruct::PropertyStruct() {}

	void PropertyStruct::Parse()
	{
		// For consistancy with previous parsings, we'll refuse to parse if asked to parse more than _MaxEntriesSmall
		// However, we may want to reconsider this choice.
		if (m_MaxEntries > _MaxEntriesSmall) return;

		DWORD dwPropCount = 0;
		for (;;)
		{
			if (dwPropCount >= m_MaxEntries) break;
			m_Props.push_back(BinToSPropValueStruct());
			if (!m_Parser.RemainingBytes()) break;
			dwPropCount++;
		}
	}

	void PropertyStruct::ParseBlocks()
	{
		auto i = 0;
		for (auto prop : m_Props)
		{
			if (i != 0)
			{
				addLine();
			}

			addHeader(L"Property[%1!d!]\r\n", i++);
			addBlock(prop.ulPropTag, L"Property = 0x%1!08X!", prop.ulPropTag.getData());

			auto propTagNames = interpretprop::PropTagToPropName(prop.ulPropTag, false);
			if (!propTagNames.bestGuess.empty())
			{
				addLine();
				addBlock(prop.ulPropTag, L"Name: %1!ws!", propTagNames.bestGuess.c_str());
			}

			if (!propTagNames.otherMatches.empty())
			{
				addLine();
				addBlock(prop.ulPropTag, L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
			}

			std::wstring PropString;
			std::wstring AltPropString;
			auto sProp = prop.getData();
			interpretprop::InterpretProp(&sProp, &PropString, &AltPropString);

			// TODO: get proper blocks here
			addLine();
			addHeader(L"PropString = %1!ws! ", strings::RemoveInvalidCharactersW(PropString, false).c_str());
			addHeader(L"AltPropString = %1!ws!", strings::RemoveInvalidCharactersW(AltPropString, false).c_str());

			auto szSmartView = InterpretPropSmartView(&sProp, nullptr, nullptr, nullptr, false, false);

			if (!szSmartView.empty())
			{
				// TODO: proper blocks here
				addLine();
				addHeader(L"Smart View: %1!ws!", szSmartView.c_str());
			}
		}
	}

	_Check_return_ SPropValueStruct PropertyStruct::BinToSPropValueStruct()
	{
		const auto ulCurrOffset = m_Parser.GetCurrentOffset();

		auto prop = SPropValueStruct{};
		prop.PropType = m_Parser.GetBlock<WORD>();
		prop.PropID = m_Parser.GetBlock<WORD>();

		prop.ulPropTag = PROP_TAG(prop.PropType, prop.PropID);
		prop.ulPropTag.setSize(prop.PropType.getSize() + prop.PropID.getSize());
		prop.ulPropTag.setOffset(prop.PropType.getOffset());
		prop.dwAlignPad = 0;

		if (m_NickName) (void) m_Parser.GetBlock<DWORD>(); // reserved

		switch (prop.PropType)
		{
		case PT_I2:
			// TODO: Insert proper property struct parsing here
			if (m_NickName) prop.Value.i = m_Parser.GetBlock<WORD>();
			if (m_NickName) m_Parser.GetBlock<WORD>();
			if (m_NickName) m_Parser.GetBlock<DWORD>();
			break;
		case PT_LONG:
			prop.Value.l = m_Parser.GetBlock<DWORD>();
			if (m_NickName) m_Parser.GetBlock<DWORD>();
			break;
		case PT_ERROR:
			prop.Value.err = m_Parser.GetBlock<DWORD>();
			if (m_NickName) m_Parser.GetBlock<DWORD>();
			break;
		case PT_R4:
			prop.Value.flt = m_Parser.GetBlock<float>();
			if (m_NickName) m_Parser.GetBlock<DWORD>();
			break;
		case PT_DOUBLE:
			prop.Value.dbl = m_Parser.GetBlock<double>();
			break;
		case PT_BOOLEAN:
			prop.Value.b = m_Parser.GetBlock<WORD>();
			if (m_NickName) m_Parser.GetBlock<WORD>();
			if (m_NickName) m_Parser.GetBlock<DWORD>();
			break;
		case PT_I8:
			prop.Value.li = m_Parser.GetBlock<LARGE_INTEGER>();
			break;
		case PT_SYSTIME:
			prop.Value.ft.dwHighDateTime = m_Parser.Get<DWORD>();
			prop.Value.ft.dwLowDateTime = m_Parser.Get<DWORD>();
			break;
		case PT_STRING8:
			if (m_NickName)
			{
				(void) m_Parser.GetBlock<LARGE_INTEGER>(); // union
				prop.Value.lpszA.cb = m_Parser.GetBlock<DWORD>();
			}
			else
			{
				prop.Value.lpszA.cb = m_Parser.GetBlock<WORD>();
			}

			prop.Value.lpszA.str = m_Parser.GetBlockStringA(prop.Value.lpszA.cb);
			break;
		case PT_BINARY:
			if (m_NickName)
			{
				(void) m_Parser.GetBlock<LARGE_INTEGER>(); // union
				prop.Value.bin.cb = m_Parser.GetBlock<DWORD>();
			}
			else
			{
				prop.Value.bin.cb = m_Parser.GetBlock<WORD>();
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			prop.Value.bin.lpb = m_Parser.GetBlockBYTES(prop.Value.bin.cb);
			break;
		case PT_UNICODE:
			if (m_NickName)
			{
				(void) m_Parser.GetBlock<LARGE_INTEGER>(); // union
				prop.Value.lpszW.cb = m_Parser.GetBlock<DWORD>();
			}
			else
			{
				prop.Value.lpszW.cb = m_Parser.GetBlock<WORD>();
			}

			prop.Value.lpszW.str = m_Parser.GetBlockStringW(prop.Value.lpszW.cb / sizeof(WCHAR));
			break;
		case PT_CLSID:
			if (m_NickName) (void) m_Parser.GetBlock<LARGE_INTEGER>(); // union
			prop.Value.lpguid = m_Parser.GetBlock<GUID>();
			break;
		case PT_MV_STRING8:
			if (m_NickName)
			{
				(void) m_Parser.GetBlock<LARGE_INTEGER>(); // union
				prop.Value.MVszA.cValues = m_Parser.GetBlock<DWORD>();
			}
			else
			{
				prop.Value.MVszA.cValues = m_Parser.GetBlock<WORD>();
			}

			if (prop.Value.MVszA.cValues)
			//if (prop.Value.MVszA.cValues && prop.Value.MVszA.cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < prop.Value.MVszA.cValues; j++)
				{
					prop.Value.MVszA.lppszA.emplace_back(m_Parser.GetBlockStringA());
				}
			}
			break;
		case PT_MV_UNICODE:
			if (m_NickName)
			{
				(void) m_Parser.GetBlock<LARGE_INTEGER>(); // union
				prop.Value.MVszW.cValues = m_Parser.GetBlock<DWORD>();
			}
			else
			{
				prop.Value.MVszW.cValues = m_Parser.GetBlock<WORD>();
			}

			if (prop.Value.MVszW.cValues)
			//if (prop.Value.MVszW.cValues && prop.Value.MVszW.cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < prop.Value.MVszW.cValues; j++)
				{
					prop.Value.MVszW.lppszW.emplace_back(m_Parser.GetBlockStringW());
				}
			}
			break;
		case PT_MV_BINARY:
			if (m_NickName)
			{
				(void) m_Parser.GetBlock<LARGE_INTEGER>(); // union
				prop.Value.MVbin.cValues = m_Parser.GetBlock<DWORD>();
			}
			else
			{
				prop.Value.MVbin.cValues = m_Parser.GetBlock<WORD>();
			}

			if (prop.Value.MVbin.cValues && prop.Value.MVbin.cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < prop.Value.MVbin.cValues; j++)
				{
					auto bin = SBinaryBlock{};
					bin.cb = m_Parser.GetBlock<DWORD>();
					// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
					bin.lpb = m_Parser.GetBlockBYTES(bin.cb);
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

	_Check_return_ std::wstring PropsToString(DWORD PropCount, LPSPropValue Prop)
	{
		std::vector<std::wstring> property;

		if (Prop)
		{
			for (DWORD i = 0; i < PropCount; i++)
			{
				std::wstring PropString;
				std::wstring AltPropString;

				property.push_back(strings::formatmessage(IDS_PROPERTYDATAHEADER, i, Prop[i].ulPropTag));

				auto propTagNames = interpretprop::PropTagToPropName(Prop[i].ulPropTag, false);
				if (!propTagNames.bestGuess.empty())
				{
					property.push_back(strings::formatmessage(IDS_PROPERTYDATANAME, propTagNames.bestGuess.c_str()));
				}

				if (!propTagNames.otherMatches.empty())
				{
					property.push_back(
						strings::formatmessage(IDS_PROPERTYDATAPARTIALMATCHES, propTagNames.otherMatches.c_str()));
				}

				interpretprop::InterpretProp(&Prop[i], &PropString, &AltPropString);
				property.push_back(strings::RemoveInvalidCharactersW(
					strings::formatmessage(IDS_PROPERTYDATA, PropString.c_str(), AltPropString.c_str()), false));

				auto szSmartView = InterpretPropSmartView(&Prop[i], nullptr, nullptr, nullptr, false, false);

				if (!szSmartView.empty())
				{
					property.push_back(strings::formatmessage(IDS_PROPERTYDATASMARTVIEW, szSmartView.c_str()));
				}
			}
		}

		return strings::join(property, L"\r\n");
	}
}