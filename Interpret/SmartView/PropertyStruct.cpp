#include <StdAfx.h>
#include <Interpret/SmartView/PropertyStruct.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/SmartView/SmartView.h>

namespace smartview
{
	PropertyStruct::PropertyStruct()
	{
		m_PropCount = 0;
		m_Prop = nullptr;
	}

	void PropertyStruct::Parse()
	{
		DWORD dwPropCount = 0;
		for (;;)
		{
			if (dwPropCount > _MaxEntriesSmall) break;
			m_Prop2.push_back(BinToSPropValueStruct(false));
			if (!m_Parser.RemainingBytes()) break;
			dwPropCount++;
		}

		m_Parser.Rewind();

		// Have to count how many properties are here.
		// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
		const auto stBookmark = m_Parser.GetCurrentOffset();

		dwPropCount = 0;

		for (;;)
		{
			const auto lpProp = BinToSPropValue(1, false);
			if (lpProp)
			{
				dwPropCount++;
			}
			else
			{
				break;
			}

			if (!m_Parser.RemainingBytes()) break;
		}

		m_Parser.SetCurrentOffset(stBookmark); // We're done with our first pass, restore the bookmark

		m_PropCount = dwPropCount;
		m_Prop = BinToSPropValue(dwPropCount, false);
	}

	_Check_return_ std::wstring PropertyStruct::ToStringInternal() { return PropsToStringBlock(m_Prop2).ToString(); }

	_Check_return_ block PropsToStringBlock(std::vector<SPropValueStruct> props)
	{
		auto blocks = block{};
		for (DWORD i = 0; i < props.size(); i++)
		{
			if (i != 0)
			{
				blocks.addLine();
			}

			blocks.addHeader(L"Property[%1!d!]\r\n", i);
			blocks.addBlock(props[i].ulPropTag, L"Property = 0x%1!08X!", props[i].ulPropTag.getData());

			auto propTagNames = interpretprop::PropTagToPropName(props[i].ulPropTag, false);
			if (!propTagNames.bestGuess.empty())
			{
				blocks.addLine();
				blocks.addBlock(props[i].ulPropTag, L"Name: %1!ws!", propTagNames.bestGuess.c_str());
			}

			if (!propTagNames.otherMatches.empty())
			{
				blocks.addLine();
				blocks.addBlock(props[i].ulPropTag, L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
			}

			std::wstring PropString;
			std::wstring AltPropString;
			auto prop = props[i].getData();
			interpretprop::InterpretProp(&prop, &PropString, &AltPropString);

			// TODO: get proper blocks here
			blocks.addLine();
			blocks.addHeader(L"PropString = %1!ws! ", PropString.c_str());
			blocks.addHeader(L"AltPropString = %1!ws!", AltPropString.c_str());

			auto szSmartView = InterpretPropSmartView(&prop, nullptr, nullptr, nullptr, false, false);

			if (!szSmartView.empty())
			{
				// TODO: proper blocks here
				blocks.addLine();
				blocks.addHeader(L"Smart View: %1!ws!", szSmartView.c_str());
			}
		}

		return blocks;
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