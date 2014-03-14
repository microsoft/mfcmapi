#pragma once
#include "Property.h"
#include <string>

Property ParseProperty(_In_ LPSPropValue lpProp);

inline std::wstring loadstring(DWORD dwID);