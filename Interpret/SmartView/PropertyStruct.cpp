#include "stdafx.h"
#include "PropertyStruct.h"
#include <Interpret/InterpretProp.h>
#include <Interpret/InterpretProp2.h>
#include <Interpret/SmartView/SmartView.h>

PropertyStruct::PropertyStruct()
{
	m_PropCount = 0;
	m_Prop = nullptr;
}

void PropertyStruct::Parse()
{
	// Have to count how many properties are here.
	// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
	auto stBookmark = m_Parser.GetCurrentOffset();

	DWORD dwPropCount = 0;

	for (;;)
	{
		auto lpProp = BinToSPropValue(1, false);
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

_Check_return_ wstring PropertyStruct::ToStringInternal()
{
	return PropsToString(m_PropCount, m_Prop);
}

_Check_return_ wstring PropsToString(DWORD PropCount, LPSPropValue Prop)
{
	wstring szProperty;

	if (Prop)
	{
		for (DWORD i = 0; i < PropCount; i++)
		{
			wstring PropString;
			wstring AltPropString;

			szProperty += formatmessage(IDS_PROPERTYDATAHEADER,
				i,
				Prop[i].ulPropTag);

			auto propTagNames = PropTagToPropName(Prop[i].ulPropTag, false);
			if (!propTagNames.bestGuess.empty())
			{
				szProperty += formatmessage(IDS_PROPERTYDATANAME,
					propTagNames.bestGuess.c_str());
			}

			if (!propTagNames.otherMatches.empty())
			{
				szProperty += formatmessage(IDS_PROPERTYDATAPARTIALMATCHES,
					propTagNames.otherMatches.c_str());
			}

			InterpretProp(&Prop[i], &PropString, &AltPropString);
			szProperty += formatmessage(IDS_PROPERTYDATA,
				PropString.c_str(),
				AltPropString.c_str());

			auto szSmartView = InterpretPropSmartView(
				&Prop[i],
				nullptr,
				nullptr,
				nullptr,
				false,
				false);

			if (!szSmartView.empty())
			{
				szProperty += formatmessage(IDS_PROPERTYDATASMARTVIEW,
					szSmartView.c_str());
			}
		}
	}

	return szProperty;
}