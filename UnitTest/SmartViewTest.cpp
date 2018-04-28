#include "stdafx.h"
#include "CppUnitTest.h"
#include "Interpret\SmartView\SmartView.h"
#include "Interpret\SmartView\SmartViewParser.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace SmartViewTest
{
	TEST_CLASS(SmartViewTest)
	{
	public:
		LPSMARTVIEWPARSER GetParser(wstring hexString, __ParsingTypeEnum iStructType)
		{
			auto bin = HexStringToBin(hexString);
			auto svp = GetSmartViewParser(iStructType, nullptr);
			if (svp)
			{
				svp->Init(bin.size(), bin.data());
				return svp;
			}

			return nullptr;
		}

		TEST_METHOD(Test_smartview)
		{
			// Set up our property arrays or nothing works
			MergeAddInArrays();

			setTestInstance(::GetModuleHandleW(L"UnitTest.dll"));
			auto additionalRenEntryIdsEx = GetParser(L"0E80320001002E0000000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F00000000013700000A80320001002E0000000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F00000000012600000B80320001002E0000000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F0000000001250000",
				IDS_STADDITIONALRENENTRYIDSEX);

			Assert::AreEqual(
				wstring(L"Additional Ren Entry IDs\r\nPersistDataCount = 3\r\n\r\nPersist Element 0:\r\nPersistID = 0x800E = 0x800E\r\nDataElementsSize = 0x0032\r\nDataElement: 0\r\n\tElementID = 0x0001 = RSF_ELID_ENTRYID\r\n\tElementDataSize = 0x002E\r\n\tElementData = cb: 46 lpb: 00000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F0000000001370000\r\n\r\nPersist Element 1:\r\nPersistID = 0x800A = 0x800A\r\nDataElementsSize = 0x0032\r\nDataElement: 0\r\n\tElementID = 0x0001 = RSF_ELID_ENTRYID\r\n\tElementDataSize = 0x002E\r\n\tElementData = cb: 46 lpb: 00000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F0000000001260000\r\n\r\nPersist Element 2:\r\nPersistID = 0x800B = 0x800B\r\nDataElementsSize = 0x0032\r\nDataElement: 0\r\n\tElementID = 0x0001 = RSF_ELID_ENTRYID\r\n\tElementDataSize = 0x002E\r\n\tElementData = cb: 46 lpb: 00000000C31A1BB1FC55D34693186631C218FEB60100CDC2D035C80A7848AA532A41B8AAE17F0000000001250000"),
				additionalRenEntryIdsEx->ToString());
		}
	};
}

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			template<> inline std::wstring ToString<__int64>(const __int64& t) { RETURN_WIDE_STRING(t); }
			template<> inline std::wstring ToString<vector<BYTE>>(const vector<BYTE>& t) { RETURN_WIDE_STRING(t.data()); }
		}
	}
}