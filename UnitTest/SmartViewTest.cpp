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

			for (auto& data : g_smartViewTestData)
			{
				auto actual = GetParser(data.hex, data.structType)->ToString();
				AreEqualEx(data.parsing, actual);

				if (data.parseAll) {
					for (ULONG i = IDS_STNOPARSING; i < IDS_STEND; i++) {
						auto parsingType = static_cast<__ParsingTypeEnum>(i);
						auto parser = GetParser(data.hex, parsingType);
						if (parser) {
							Logger::WriteMessage(format(L"Testing %ws\n", AddInStructTypeToString(parsingType).c_str()).c_str());
							Assert::IsTrue(parser->ToString().length() != 0);
						}
					}
				}
			}
		}
	};
}