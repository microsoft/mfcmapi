#pragma once
#include "Property.h"

Property ParseProperty(_In_ LPSPropValue lpProp);
wstring GUIDToWstring(_In_opt_ LPCGUID lpGUID);
wstring GUIDToWstringAndName(_In_opt_ LPCGUID lpGUID);
wstring BinToTextString(_In_ LPSBinary lpBin, bool bMultiLine);
wstring BinToHexString(_In_opt_ LPSBinary lpBin, bool bPrependCB);
void FileTimeToString(_In_ FILETIME* lpFileTime, _In_ wstring& PropString, _In_opt_ wstring& AltPropString);