#include "stdafx.h"
#include "CppUnitTest.h"
#include "Interpret\SmartView\SmartView.h"
#include "Interpret\SmartView\SmartViewParser.h"
#include "SmartViewTestData.h"
#include "MFCMAPI.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace SmartViewTest
{
	TEST_CLASS(SmartViewTest)
	{
	private:
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

		TEST_METHOD(Test_smartview)
		{
			auto testData = SmartViewTestData::getTestData(::GetModuleHandleW(L"UnitTest.dll"));
			test(testData);
		}
	};
}