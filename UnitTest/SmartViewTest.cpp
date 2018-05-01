#include "stdafx.h"
#include "CppUnitTest.h"
#include "Interpret\SmartView\SmartView.h"
#include "Interpret\SmartView\SmartViewParser.h"
#include "SmartViewTestData.h"

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
				Assert::AreEqual(
					data.parsing,
					GetParser(data.hex, data.structType)->ToString());
			}
		}
	};
}