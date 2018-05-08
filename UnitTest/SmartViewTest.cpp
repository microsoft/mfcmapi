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
			for (auto data : testData)
			{
				auto parser = GetParser(data.hex, data.structType);
				if (parser != nullptr)
				{
					AreEqualEx(data.expected, parser->ToString(), data.testName.c_str());
				}

				if (data.parseAll) {
					for (ULONG i = IDS_STNOPARSING; i < IDS_STEND; i++) {
						auto parsingType = static_cast<__ParsingTypeEnum>(i);
						parser = GetParser(data.hex, parsingType);
						if (parser) {
							//Logger::WriteMessage(format(L"Testing %ws\n", AddInStructTypeToString(parsingType).c_str()).c_str());
							Assert::IsTrue(parser->ToString().length() != 0);
						}
					}
				}
			}
		}
	};
}