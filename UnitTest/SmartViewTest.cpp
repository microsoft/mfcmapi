#include "stdafx.h"
#include "CppUnitTest.h"
#include "Interpret\SmartView\SmartView.h"
#include "Interpret\SmartView\SmartViewParser.h"
#include "SmartViewTestData.h"
#include "MFCMAPI.h"
#include "resource.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace SmartViewTest
{
	TEST_CLASS(SmartViewTest)
	{
	private:
		const bool parseAll = false;
		wstring ParseString(wstring hexString, __ParsingTypeEnum iStructType)
		{
			auto bin = HexStringToBin(hexString);
			return InterpretBinaryAsString({ bin.size(), bin.data() }, iStructType, nullptr);
		}

		void test(vector<SmartViewTestData::SmartViewTestData> testData)
		{
			for (auto data : testData)
			{
				AreEqualEx(data.expected, ParseString(data.hex, data.structType), data.testName.c_str());

				if (data.parseAll) {
					for (ULONG i = IDS_STNOPARSING; i < IDS_STEND; i++) {
						//Logger::WriteMessage(format(L"Testing %ws\n", AddInStructTypeToString(static_cast<__ParsingTypeEnum>(i)).c_str()).c_str());
						Assert::IsTrue(ParseString(data.hex, data.structType).length() != 0);
					}
				}
			}
		}

	public:
		TEST_CLASS_INITIALIZE(Initialize_smartview)
		{
			// Set up our property arrays or nothing works
			MergeAddInArrays();

			RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD = 1;
			RegKeys[regkeyUSE_GETPROPLIST].ulCurDWORD = 1;
			RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD = 1;
			RegKeys[regkeyCACHE_NAME_DPROPS].ulCurDWORD = 1;

			setTestInstance(::GetModuleHandleW(L"UnitTest.dll"));
		}

		TEST_METHOD(Test_STADDITIONALRENENTRYIDSEX)
		{
			test(SmartViewTestData::loadTestData(::GetModuleHandleW(L"UnitTest.dll"),
				{
					SmartViewTestData::SmartViewTestResource{ IDS_STADDITIONALRENENTRYIDSEX, parseAll, IDR_SV1AEI1IN, IDR_SV1AEI1OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STADDITIONALRENENTRYIDSEX, parseAll, IDR_SV1AEI2IN, IDR_SV1AEI2OUT },
				}));
		}

		TEST_METHOD(Test_STAPPOINTMENTRECURRENCEPATTERN)
		{
			test(SmartViewTestData::loadTestData(::GetModuleHandleW(L"UnitTest.dll"),
				{
					SmartViewTestData::SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP1IN, IDR_SV2ARP1OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP2IN, IDR_SV2ARP2OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP3IN, IDR_SV2ARP3OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parseAll, IDR_SV2ARP4IN, IDR_SV2ARP4OUT },
				}));
		}

		TEST_METHOD(Test_STCONVERSATIONINDEX)
		{
			test(SmartViewTestData::loadTestData(::GetModuleHandleW(L"UnitTest.dll"),
				{
					SmartViewTestData::SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parseAll, IDR_SV3CI1IN, IDR_SV3CI1OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parseAll, IDR_SV3CI2IN, IDR_SV3CI2OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parseAll, IDR_SV3CI3IN, IDR_SV3CI3OUT },
				}));
		}

		TEST_METHOD(Test_STENTRYID)
		{
			test(SmartViewTestData::loadTestData(::GetModuleHandleW(L"UnitTest.dll"),
				{
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID1IN, IDR_SV4EID1OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID2IN, IDR_SV4EID2OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID3IN, IDR_SV4EID3OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID4IN, IDR_SV4EID4OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID5IN, IDR_SV4EID5OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID6IN, IDR_SV4EID6OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID7IN, IDR_SV4EID7OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID8IN, IDR_SV4EID8OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID9IN, IDR_SV4EID9OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID10IN, IDR_SV4EID10OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID11IN, IDR_SV4EID11OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID12IN, IDR_SV4EID12OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID13IN, IDR_SV4EID13OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID14IN, IDR_SV4EID14OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID15IN, IDR_SV4EID15OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID16IN, IDR_SV4EID16OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID17IN, IDR_SV4EID17OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID18IN, IDR_SV4EID18OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID19IN, IDR_SV4EID19OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID20IN, IDR_SV4EID20OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID21IN, IDR_SV4EID21OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID22IN, IDR_SV4EID22OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID23IN, IDR_SV4EID23OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID24IN, IDR_SV4EID24OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID25IN, IDR_SV4EID25OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID26IN, IDR_SV4EID26OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID27IN, IDR_SV4EID27OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID28IN, IDR_SV4EID28OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID29IN, IDR_SV4EID29OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID30IN, IDR_SV4EID30OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID31IN, IDR_SV4EID31OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID32IN, IDR_SV4EID32OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID33IN, IDR_SV4EID33OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID34IN, IDR_SV4EID34OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID35IN, IDR_SV4EID35OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID36IN, IDR_SV4EID36OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID37IN, IDR_SV4EID37OUT },
					SmartViewTestData::SmartViewTestResource{ IDS_STENTRYID, parseAll, IDR_SV4EID38IN, IDR_SV4EID38OUT }
				}));
		}
	};
}