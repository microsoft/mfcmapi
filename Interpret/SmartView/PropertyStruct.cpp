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
		// Have to count how many properties are here.
		// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
		const auto stBookmark = m_Parser.GetCurrentOffset();

		DWORD dwPropCount = 0;

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

	_Check_return_ std::wstring PropertyStruct::ToStringInternal()
	{
		return PropsToString(m_PropCount, m_Prop);
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

				property.push_back(strings::formatmessage(IDS_PROPERTYDATAHEADER,
					i,
					Prop[i].ulPropTag));

				auto propTagNames = interpretprop::PropTagToPropName(Prop[i].ulPropTag, false);
				if (!propTagNames.bestGuess.empty())
				{
					property.push_back(strings::formatmessage(IDS_PROPERTYDATANAME,
						propTagNames.bestGuess.c_str()));
				}

				if (!propTagNames.otherMatches.empty())
				{
					property.push_back(strings::formatmessage(IDS_PROPERTYDATAPARTIALMATCHES,
						propTagNames.otherMatches.c_str()));
				}

				interpretprop::InterpretProp(&Prop[i], &PropString, &AltPropString);
				property.push_back(strings::RemoveInvalidCharactersW(strings::formatmessage(IDS_PROPERTYDATA,
					PropString.c_str(),
					AltPropString.c_str()), false));

				auto szSmartView = InterpretPropSmartView(
					&Prop[i],
					nullptr,
					nullptr,
					nullptr,
					false,
					false);

				if (!szSmartView.empty())
				{
					property.push_back(strings::formatmessage(IDS_PROPERTYDATASMARTVIEW,
						szSmartView.c_str()));
				}
			}
		}

		return strings::join(property, L"\r\n");
	}
}