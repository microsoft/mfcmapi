#include "stdafx.h"
#include "SmartViewTestData.h"
#include "resource.h"

namespace SmartViewTestData
{
	struct SmartViewTestResource {
		__ParsingTypeEnum structType;
		bool parseAll;
		WORD hex;
		WORD expected;
	};

	wstring loadfile(HMODULE handle, int name)
	{
		auto rc = ::FindResource(handle, MAKEINTRESOURCE(name), MAKEINTRESOURCE(TEXTFILE));
		auto rcData = ::LoadResource(handle, rc);
		auto size = ::SizeofResource(handle, rc);
		auto data = static_cast<const char*>(::LockResource(rcData));
		auto ansi = std::string(data, size);
		return wstring(ansi.begin(), ansi.end());
	}

	vector<SmartViewTestData> getTestData(HMODULE handle)
	{
		auto resources = {
			SmartViewTestResource{IDS_STADDITIONALRENENTRYIDSEX, true, IDR_SV1AEI1IN, IDR_SV1AEI1OUT},
			SmartViewTestResource{IDS_STADDITIONALRENENTRYIDSEX, true, IDR_SV1AEI2IN, IDR_SV1AEI2OUT},

			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, true, IDR_SV2ARP1IN, IDR_SV2ARP1OUT },
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, true, IDR_SV2ARP2IN, IDR_SV2ARP2OUT },
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, true, IDR_SV2ARP3IN, IDR_SV2ARP3OUT },
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, true, IDR_SV2ARP4IN, IDR_SV2ARP4OUT },
		};

		vector<SmartViewTestData> testData;
		for (auto resource : resources)
		{
			testData.push_back(SmartViewTestData
			{
				resource.structType, resource.parseAll,
				loadfile(handle, resource.hex),
				loadfile(handle, resource.expected)
			});
		}

		return testData;
	}
}