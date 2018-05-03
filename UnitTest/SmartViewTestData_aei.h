#pragma once
#include "SmartViewTestData.h"

namespace SmartViewTestData
{
	class SmartViewTestData_aei
	{
	public:
		static void init()
		{
			auto aei1 = SmartViewTestData
			{
				IDS_STADDITIONALRENENTRYIDSEX, true,
				L"0E80320001002E0000000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA5"
				L"32A41B8AAE17F00000000013700000A80320001002E0000000000C31A1BB1FC55D34693186631C2"
				L"18FEB60100CDC2D035C80A7848AA532A41B8AAE17F00000000012600000B80320001002E0000000"
				L"000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F00000000"
				L"01250000",
				L"Additional Ren Entry IDs\r\n"
				L"PersistDataCount = 3\r\n"
				L"\r\n"
				L"Persist Element 0:\r\n"
				L"PersistID = 0x800E = 0x800E\r\n"
				L"DataElementsSize = 0x0032\r\n"
				L"DataElement: 0\r\n"
				L"\tElementID = 0x0001 = RSF_ELID_ENTRYID\r\n"
				L"\tElementDataSize = 0x002E\r\n"
				L"\tElementData = cb: 46 lpb: 00000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F0000000001370000\r\n"
				L"\r\n"
				L"Persist Element 1:\r\n"
				L"PersistID = 0x800A = 0x800A\r\n"
				L"DataElementsSize = 0x0032\r\n"
				L"DataElement: 0\r\n"
				L"\tElementID = 0x0001 = RSF_ELID_ENTRYID\r\n"
				L"\tElementDataSize = 0x002E\r\n"
				L"\tElementData = cb: 46 lpb: 00000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F0000000001260000\r\n"
				L"\r\n"
				L"Persist Element 2:\r\n"
				L"PersistID = 0x800B = 0x800B\r\n"
				L"DataElementsSize = 0x0032\r\n"
				L"DataElement: 0\r\n"
				L"\tElementID = 0x0001 = RSF_ELID_ENTRYID\r\n"
				L"\tElementDataSize = 0x002E\r\n"
				L"\tElementData = cb: 46 lpb: 00000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F0000000001250000"
			};

			g_smartViewTestData.push_back(aei1);
		}
	};
}