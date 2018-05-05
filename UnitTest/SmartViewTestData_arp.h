#pragma once
#include "SmartViewTestData.h"
#include "resource.h"

namespace SmartViewTestData
{
	class SmartViewTestData_arp
	{
	public:
		static void init(HMODULE handle)
		{
			auto arp1 = SmartViewTestData
			{
				IDS_STAPPOINTMENTRECURRENCEPATTERN, true,
				loadfile(handle, IDR_SV2ARP1IN),
				loadfile(handle, IDR_SV2ARP1OUT)
			};
			auto arp2 = SmartViewTestData
			{
				IDS_STAPPOINTMENTRECURRENCEPATTERN, false,
				loadfile(handle, IDR_SV2ARP2IN),
				loadfile(handle, IDR_SV2ARP2OUT)
			};
			auto arp3 = SmartViewTestData
			{
				IDS_STAPPOINTMENTRECURRENCEPATTERN, false,
				loadfile(handle, IDR_SV2ARP3IN),
				loadfile(handle, IDR_SV2ARP3OUT)
			};
			auto arp4 = SmartViewTestData
			{
				IDS_STAPPOINTMENTRECURRENCEPATTERN, false,
				loadfile(handle, IDR_SV2ARP4IN),
				loadfile(handle, IDR_SV2ARP4OUT)
			};

			g_smartViewTestData.push_back(arp1);
			g_smartViewTestData.push_back(arp2);
			g_smartViewTestData.push_back(arp3);
			g_smartViewTestData.push_back(arp4);
		}
	};
}