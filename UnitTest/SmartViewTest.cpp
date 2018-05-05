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
			auto handle = ::GetModuleHandleW(L"UnitTest.dll");
			setTestInstance(handle);
			SmartViewTestData::init(handle);
		}

		TEST_METHOD(Test_smartview)
		{
			for (auto& data : SmartViewTestData::g_smartViewTestData)
			{
				auto actual = GetParser(data.hex, data.structType)->ToString();
				AreEqualEx(data.expected, actual);

				if (data.parseAll) {
					for (ULONG i = IDS_STNOPARSING; i < IDS_STEND; i++) {
						auto parsingType = static_cast<__ParsingTypeEnum>(i);
						auto parser = GetParser(data.hex, parsingType);
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