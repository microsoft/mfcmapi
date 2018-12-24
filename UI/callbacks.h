#pragma once
#include <MAPI/MAPIFunctions.h>

namespace ui
{
	namespace callbacks
	{
		void init();

		LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB);
	} // namespace callbacks
} // namespace ui