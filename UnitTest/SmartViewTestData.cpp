#include "stdafx.h"
#include "SmartViewTestData.h"
#include "resource.h"

namespace SmartViewTestData
{
	struct SmartViewTestResource {
		__ParsingTypeEnum structType;
		bool parseAll;
		DWORD hex;
		DWORD expected;
	};

	// Resource files saved in unicode have a byte order mark of 0xfeff
	// We load these in and strip the BOM.
	// Otherwise we load as ansi and convert to unicode
	wstring loadfile(HMODULE handle, int name)
	{
		auto rc = ::FindResource(handle, MAKEINTRESOURCE(name), MAKEINTRESOURCE(TEXTFILE));
		auto rcData = ::LoadResource(handle, rc);
		auto cb = ::SizeofResource(handle, rc);
		auto bytes = ::LockResource(rcData);
		auto data = static_cast<const BYTE*>(bytes);
		if (cb >= 2 && data[0] == 0xfe && data[1] == 0xff)
		{
			// Skip the byte order mark
			auto wstr = static_cast<const wchar_t*>(bytes);
			auto cch = cb / sizeof(wchar_t);
			return std::wstring(wstr + 1, cch - 1);
		}

		auto str = std::string(static_cast<const char*>(bytes), cb);
		return stringTowstring(str);

	}

	vector<SmartViewTestData> getTestData(HMODULE handle)
	{
		const bool parseAll = false;
		auto resources = {
			SmartViewTestResource{ IDS_STADDITIONALRENENTRYIDSEX, parseAll, IDR_SV1AEI1IN, IDR_SV1AEI1OUT },
			SmartViewTestResource{ IDS_STADDITIONALRENENTRYIDSEX, parseAll, IDR_SV1AEI2IN, IDR_SV1AEI2OUT },

			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP1IN, IDR_SV2ARP1OUT },
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP2IN, IDR_SV2ARP2OUT },
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP3IN, IDR_SV2ARP3OUT },
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP4IN, IDR_SV2ARP4OUT },

			SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parseAll, IDR_SV3CI1IN, IDR_SV3CI1OUT },
			SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parseAll, IDR_SV3CI2IN, IDR_SV3CI2OUT },
			SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parseAll, IDR_SV3CI3IN, IDR_SV3CI3OUT },
		};

		vector<SmartViewTestData> testData;
		for (auto resource : resources)
		{
			testData.push_back(SmartViewTestData
				{
					resource.structType, resource.parseAll,
					format(L"%d/%d", resource.hex, resource.expected),
					loadfile(handle, resource.hex),
					loadfile(handle, resource.expected)
				});
		}

		return testData;
	}
}