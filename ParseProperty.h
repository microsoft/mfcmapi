#pragma once
#include "Property.h"

Property ParseProperty(_In_ LPSPropValue lpProp);
_Check_return_ wstring GUIDToWstring(_In_opt_ LPCGUID lpGUID);
_Check_return_ wstring GUIDToWstringAndName(_In_opt_ LPCGUID lpGUID);