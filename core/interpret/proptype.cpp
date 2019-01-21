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
} // namespace proptype