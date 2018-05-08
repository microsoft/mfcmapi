#include "stdafx.h"
#include "SmartViewTestData.h"
#include "resource.h"

namespace SmartViewTestData
{
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

	vector<SmartViewTestData> loadTestData(HMODULE handle, std::initializer_list<SmartViewTestResource> resources)
	{
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

	vector<SmartViewTestData> getTestData(HMODULE handle)
	{
		const bool parseAll = false;
		auto resources = {
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP1IN, IDR_SV2ARP1OUT },
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP2IN, IDR_SV2ARP2OUT },
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP3IN, IDR_SV2ARP3OUT },
			SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP4IN, IDR_SV2ARP4OUT },

			SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parseAll, IDR_SV3CI1IN, IDR_SV3CI1OUT },
			SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parseAll, IDR_SV3CI2IN, IDR_SV3CI2OUT },
			SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parseAll, IDR_SV3CI3IN, IDR_SV3CI3OUT },

			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID1IN, IDR_SV4EID1OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID2IN, IDR_SV4EID2OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID3IN, IDR_SV4EID3OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID4IN, IDR_SV4EID4OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID5IN, IDR_SV4EID5OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID6IN, IDR_SV4EID6OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID7IN, IDR_SV4EID7OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID8IN, IDR_SV4EID8OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID9IN, IDR_SV4EID9OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID10IN, IDR_SV4EID10OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID11IN, IDR_SV4EID11OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID12IN, IDR_SV4EID12OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID13IN, IDR_SV4EID13OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID14IN, IDR_SV4EID14OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID15IN, IDR_SV4EID15OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID16IN, IDR_SV4EID16OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID17IN, IDR_SV4EID17OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID18IN, IDR_SV4EID18OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID19IN, IDR_SV4EID19OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID20IN, IDR_SV4EID20OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID21IN, IDR_SV4EID21OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID22IN, IDR_SV4EID22OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID23IN, IDR_SV4EID23OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID24IN, IDR_SV4EID24OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID25IN, IDR_SV4EID25OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID26IN, IDR_SV4EID26OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID27IN, IDR_SV4EID27OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID28IN, IDR_SV4EID28OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID29IN, IDR_SV4EID29OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID30IN, IDR_SV4EID30OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID31IN, IDR_SV4EID31OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID32IN, IDR_SV4EID32OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID33IN, IDR_SV4EID33OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID34IN, IDR_SV4EID34OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID35IN, IDR_SV4EID35OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID36IN, IDR_SV4EID36OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID37IN, IDR_SV4EID37OUT },
			SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID38IN, IDR_SV4EID38OUT },
		};

		return loadTestData(handle, resources);
	}
}