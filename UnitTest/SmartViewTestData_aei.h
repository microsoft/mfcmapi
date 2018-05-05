#pragma once
#include "SmartViewTestData.h"
#include "resource.h"

namespace SmartViewTestData
{
	class SmartViewTestData_aei
	{
	public:
		static void init(HMODULE handle)
		{
			auto aei1 = SmartViewTestData
			{
				IDS_STADDITIONALRENENTRYIDSEX, true,
				loadfile(handle, IDR_SV1AEI1IN),
				loadfile(handle, IDR_SV1AEI1OUT)
			};
			auto aei2 = SmartViewTestData
			{
				IDS_STADDITIONALRENENTRYIDSEX, true,
				loadfile(handle, IDR_SV1AEI2IN),
				loadfile(handle, IDR_SV1AEI2OUT)
			};

			g_smartViewTestData.push_back(aei1);
			g_smartViewTestData.push_back(aei2);
		}
	};
}