#include <core/stdafx.h>
#include <core/interpret/proptype.h>
#include <core/utility/strings.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>

namespace proptype
{
	std::wstring TypeToString(ULONG ulPropTag)
	{
		std::wstring tmpPropType;

		auto bNeedInstance = false;
		if (ulPropTag & MV_INSTANCE)
		{
			ulPropTag &= ~MV_INSTANCE;
			bNeedInstance = true;
		}

		auto bTypeFound = false;

		for (const auto& propType : PropTypeArray)
		{
			if (propType.ulValue == PROP_TYPE(ulPropTag))
			{
				tmpPropType = propType.lpszName;
				bTypeFound = true;
				break;
			}
		}

		if (!bTypeFound) tmpPropType = strings::format(L"0x%04x", PROP_TYPE(ulPropTag)); // STRING_OK

		if (bNeedInstance) tmpPropType += L" | MV_INSTANCE"; // STRING_OK
		return tmpPropType;
	}

	_Check_return_ ULONG PropTypeNameToPropType(_In_ const std::wstring& lpszPropType)
	{
		if (lpszPropType.empty() || PropTypeArray.empty()) return PT_UNSPECIFIED;

		// Check for numbers first before trying the string as an array lookup.
		// This will translate '0x102' to 0x102, 0x3 to 3, etc.
		const auto ulType = strings::wstringToUlong(lpszPropType, 16);
		if (ulType != NULL) return ulType;

		auto ulPropType = PT_UNSPECIFIED;

		for (const auto& propType : PropTypeArray)
		{
			if (0 == lstrcmpiW(lpszPropType.c_str(), propType.lpszName))
			{
				ulPropType = propType.ulValue;
				break;
			}
		}

		return ulPropType;
	}
} // namespace proptype