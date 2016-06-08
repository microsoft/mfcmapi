#pragma once
#include "Property.h"

Property ParseProperty(_In_ LPSPropValue lpProp);
wstring BinToTextString(_In_ LPSBinary lpBin, bool bMultiLine);
wstring BinToHexString(_In_opt_count_(cb) LPBYTE lpb, size_t cb, bool bPrependCB);
wstring BinToHexString(_In_opt_ LPSBinary lpBin, bool bPrependCB);