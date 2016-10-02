#include "stdafx.h"
#include "PropertyStruct.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"

PropertyStruct::PropertyStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_PropCount = 0;
	m_Prop = nullptr;
}

PropertyStruct::~PropertyStruct()
{
	DeleteSPropVal(m_PropCount, m_Prop);
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
			DeleteSPropVal(1, lpProp);
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

			wstring szExactMatches;
			wstring szPartialMatches;
			PropTagToPropName(Prop[i].ulPropTag, false, szExactMatches, szPartialMatches);
			if (!szExactMatches.empty())
			{
				szProperty += formatmessage(IDS_PROPERTYDATAEXACTMATCHES,
					szExactMatches.c_str());
			}

			if (!szPartialMatches.empty())
			{
				szProperty += formatmessage(IDS_PROPERTYDATAPARTIALMATCHES,
					szPartialMatches.c_str());
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

void DeleteSPropVal(ULONG cVal, _In_count_(cVal) LPSPropValue lpsPropVal)
{
	if (!lpsPropVal) return;
	for (ULONG i = 0; i < cVal; i++)
	{
		switch (PROP_TYPE(lpsPropVal[i].ulPropTag))
		{
		case PT_UNICODE:
			delete[] lpsPropVal[i].Value.lpszW;
			break;
		case PT_STRING8:
			delete[] lpsPropVal[i].Value.lpszA;
			break;
		case PT_BINARY:
			delete[] lpsPropVal[i].Value.bin.lpb;
			break;
		case PT_MV_STRING8:
			if (lpsPropVal[i].Value.MVszA.lppszA)
			{
				for (ULONG j = 0; j < lpsPropVal[i].Value.MVszA.cValues; j++)
				{
					delete[] lpsPropVal[i].Value.MVszA.lppszA[j];
				}

				delete[] lpsPropVal[i].Value.MVszA.lppszA;
			}
			break;
		case PT_MV_UNICODE:
			if (lpsPropVal[i].Value.MVszW.lppszW)
			{
				for (ULONG j = 0; j < lpsPropVal[i].Value.MVszW.cValues; j++)
				{
					delete[] lpsPropVal[i].Value.MVszW.lppszW[j];
				}

				delete[] lpsPropVal[i].Value.MVszW.lppszW;
			}
			break;
		case PT_MV_BINARY:
			if (lpsPropVal[i].Value.MVbin.lpbin)
			{
				for (ULONG j = 0; j < lpsPropVal[i].Value.MVbin.cValues; j++)
				{
					delete[] lpsPropVal[i].Value.MVbin.lpbin[j].lpb;
				}

				delete[] lpsPropVal[i].Value.MVbin.lpbin;
			}
			break;
		}
	}

	delete[] lpsPropVal;
}