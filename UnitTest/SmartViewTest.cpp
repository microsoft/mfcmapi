#include "stdafx.h"
#include "CppUnitTest.h"
#include "Interpret/SmartView/SmartView.h"
#include "MFCMAPI.h"
#include "resource.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace SmartViewTest
{
	TEST_CLASS(SmartViewTest)
	{
	private:
		struct SmartViewTestResource
		{
			__ParsingTypeEnum structType;
			bool parseAll;
			DWORD hex;
			DWORD expected;
		};

		struct SmartViewTestData
		{
			__ParsingTypeEnum structType;
			bool parseAll;
			wstring testName;
			vector<BYTE> hex;
			wstring expected;
		};

		static const bool parse_all = false;
		static const bool assert_on_failure = true;
		static const bool limit_output = false;

		// Assert::AreEqual doesn't do full logging, so we roll our own
		void AreEqualEx(const wstring& expected, const wstring& actual, const wchar_t* message = nullptr, const __LineInfo* pLineInfo = nullptr) const
		{
			if (expected != actual)
			{
				if (message != nullptr)
				{
					Logger::WriteMessage(format(L"Test: %ws\n", message).c_str());
				}

				Logger::WriteMessage(L"Diff:\n");

				auto splitExpected = split(expected, L'\n');
				auto splitActual = split(actual, L'\n');
				auto errorCount = 0;
				for (size_t line = 0; line < splitExpected.size() && line < splitActual.size() && (errorCount < 4 || !limit_output); line++)
				{
					if (splitExpected[line] != splitActual[line])
					{
						errorCount++;
						Logger::WriteMessage(format(L"[%d]\n\"%ws\"\n\"%ws\"\n", line + 1, splitExpected[line].c_str(), splitActual[line].c_str()).c_str());
						auto lineErrorCount = 0;
						for (size_t ch = 0; ch < splitExpected[line].size() && ch < splitActual[line].size() && (lineErrorCount < 10 || !limit_output); ch++)
						{
							wchar_t expectedChar = splitExpected[line][ch];
							wchar_t actualChar = splitActual[line][ch];
							if (expectedChar != actualChar)
							{
								lineErrorCount++;
								Logger::WriteMessage(format(L"[%d]: %X (%wc) != %X (%wc)\n", ch + 1, expectedChar, expectedChar, actualChar, actualChar).c_str());
							}
						}
					}
				}

				if (!limit_output)
				{
					Logger::WriteMessage(L"\n");
					Logger::WriteMessage(format(L"Expected:\n\"%ws\"\n\n", expected.c_str()).c_str());
					Logger::WriteMessage(format(L"Actual:\n\"%ws\"", actual.c_str()).c_str());
				}

				if (assert_on_failure)
				{
					Assert::Fail(ToString(message).c_str(), pLineInfo);
				}
			}
		}

		void test(vector<SmartViewTestData> testData) const
		{
			for (auto data : testData)
			{
				auto actual = InterpretBinaryAsString({ static_cast<ULONG>(data.hex.size()), data.hex.data() }, data.structType, nullptr);
				AreEqualEx(data.expected, actual, data.testName.c_str());

				if (data.parseAll)
				{
					for (ULONG i = IDS_STNOPARSING; i < IDS_STEND; i++)
					{
						const auto structType = static_cast<__ParsingTypeEnum>(i);
						actual = InterpretBinaryAsString({ static_cast<ULONG>(data.hex.size()), data.hex.data() }, structType, nullptr);
						//Logger::WriteMessage(format(L"Testing %ws\n", AddInStructTypeToString(structType).c_str()).c_str());
						Assert::IsTrue(actual.length() != 0);
					}
				}
			}
		}

		// Resource files saved in unicode have a byte order mark of 0xfeff
		// We load these in and strip the BOM.
		// Otherwise we load as ansi and convert to unicode
		static wstring loadfile(const HMODULE handle, const int name)
		{
			const auto rc = ::FindResource(handle, MAKEINTRESOURCE(name), MAKEINTRESOURCE(TEXTFILE));
			const auto rcData = LoadResource(handle, rc);
			const auto cb = SizeofResource(handle, rc);
			const auto bytes = LockResource(rcData);
			const auto data = static_cast<const BYTE*>(bytes);
			if (cb >= 2 && data[0] == 0xfe && data[1] == 0xff)
			{
				// Skip the byte order mark
				const auto wstr = static_cast<const wchar_t*>(bytes);
				const auto cch = cb / sizeof(wchar_t);
				return std::wstring(wstr + 1, cch - 1);
			}

			const auto str = std::string(static_cast<const char*>(bytes), cb);
			return stringTowstring(str);
		}

		static vector<SmartViewTestData> loadTestData(std::initializer_list<SmartViewTestResource> resources)
		{
			static auto handle = GetModuleHandleW(L"UnitTest.dll");
			vector<SmartViewTestData> testData;
			for (auto resource : resources)
			{
				testData.push_back(SmartViewTestData
					{
						resource.structType,
						resource.parseAll,
						format(L"%d/%d", resource.hex, resource.expected),
						HexStringToBin(loadfile(handle, resource.hex)),
						loadfile(handle, resource.expected)
					});
			}

			return testData;
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

			setTestInstance(GetModuleHandleW(L"UnitTest.dll"));
		}

		TEST_METHOD(Test_STADDITIONALRENENTRYIDSEX)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STADDITIONALRENENTRYIDSEX, parse_all, IDR_SV1AEI1IN, IDR_SV1AEI1OUT },
				SmartViewTestResource{ IDS_STADDITIONALRENENTRYIDSEX, parse_all, IDR_SV1AEI2IN, IDR_SV1AEI2OUT },
				}));
		}

		TEST_METHOD(Test_STAPPOINTMENTRECURRENCEPATTERN)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parse_all, IDR_SV2ARP1IN, IDR_SV2ARP1OUT },
				SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parse_all, IDR_SV2ARP2IN, IDR_SV2ARP2OUT },
				SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parse_all, IDR_SV2ARP3IN, IDR_SV2ARP3OUT },
				SmartViewTestResource{ IDS_STAPPOINTMENTRECURRENCEPATTERN, parse_all, IDR_SV2ARP4IN, IDR_SV2ARP4OUT },
				}));
		}

		TEST_METHOD(Test_STCONVERSATIONINDEX)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parse_all, IDR_SV3CI1IN, IDR_SV3CI1OUT },
				SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parse_all, IDR_SV3CI2IN, IDR_SV3CI2OUT },
				SmartViewTestResource{ IDS_STCONVERSATIONINDEX, parse_all, IDR_SV3CI3IN, IDR_SV3CI3OUT },
				}));
		}

		TEST_METHOD(Test_STENTRYID)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID1IN, IDR_SV4EID1OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID2IN, IDR_SV4EID2OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID3IN, IDR_SV4EID3OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID4IN, IDR_SV4EID4OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID5IN, IDR_SV4EID5OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID6IN, IDR_SV4EID6OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID7IN, IDR_SV4EID7OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID8IN, IDR_SV4EID8OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID9IN, IDR_SV4EID9OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID10IN, IDR_SV4EID10OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID11IN, IDR_SV4EID11OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID12IN, IDR_SV4EID12OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID13IN, IDR_SV4EID13OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID14IN, IDR_SV4EID14OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID15IN, IDR_SV4EID15OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID16IN, IDR_SV4EID16OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID17IN, IDR_SV4EID17OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID18IN, IDR_SV4EID18OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID19IN, IDR_SV4EID19OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID20IN, IDR_SV4EID20OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID21IN, IDR_SV4EID21OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID22IN, IDR_SV4EID22OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID23IN, IDR_SV4EID23OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID24IN, IDR_SV4EID24OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID25IN, IDR_SV4EID25OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID26IN, IDR_SV4EID26OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID27IN, IDR_SV4EID27OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID28IN, IDR_SV4EID28OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID29IN, IDR_SV4EID29OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID30IN, IDR_SV4EID30OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID31IN, IDR_SV4EID31OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID32IN, IDR_SV4EID32OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID33IN, IDR_SV4EID33OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID34IN, IDR_SV4EID34OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID35IN, IDR_SV4EID35OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID36IN, IDR_SV4EID36OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID37IN, IDR_SV4EID37OUT },
				SmartViewTestResource{ IDS_STENTRYID, parse_all, IDR_SV4EID38IN, IDR_SV4EID38OUT }
				}));
		}

		TEST_METHOD(Test_STENTRYLIST)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STENTRYLIST, parse_all, IDR_SV5EL1IN, IDR_SV5EL1OUT },
				SmartViewTestResource{ IDS_STENTRYLIST, parse_all, IDR_SV5EL2IN, IDR_SV5EL2OUT },
				}));
		}

		TEST_METHOD(Test_STEXTENDEDFOLDERFLAGS)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STEXTENDEDFOLDERFLAGS, parse_all, IDR_SV6EFF1IN, IDR_SV6EFF1OUT },
				SmartViewTestResource{ IDS_STEXTENDEDFOLDERFLAGS, parse_all, IDR_SV6EFF2IN, IDR_SV6EFF2OUT },
				}));
		}

		TEST_METHOD(Test_STEXTENDEDRULECONDITION)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STEXTENDEDRULECONDITION, parse_all, IDR_SV7EXRULE1IN, IDR_SV7EXRULE1OUT },
				}));
		}

		TEST_METHOD(Test_STFLATENTRYLIST)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STFLATENTRYLIST, parse_all, IDR_SV8FE1IN, IDR_SV8FE1OUT },
				SmartViewTestResource{ IDS_STFLATENTRYLIST, parse_all, IDR_SV8FE2IN, IDR_SV8FE2OUT },
				}));
		}

		TEST_METHOD(Test_STFOLDERUSERFIELDS)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STFOLDERUSERFIELDS, parse_all, IDR_SV9FUF1IN, IDR_SV9FUF1OUT },
				SmartViewTestResource{ IDS_STFOLDERUSERFIELDS, parse_all, IDR_SV9FUF2IN, IDR_SV9FUF2OUT },
				}));
		}

		TEST_METHOD(Test_STGLOBALOBJECTID)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STGLOBALOBJECTID, parse_all, IDR_SV10GOID1IN, IDR_SV10GOID1OUT },
				SmartViewTestResource{ IDS_STGLOBALOBJECTID, parse_all, IDR_SV10GOID2IN, IDR_SV10GOID2OUT },
				}));
		}

		TEST_METHOD(Test_STPROPERTY)
		{
			test(loadTestData({
				//SmartViewTestResource{ IDS_STPROPERTY, parse_all, IDR_SV11PROP1IN, IDR_SV11PROP1OUT },
				SmartViewTestResource{ IDS_STPROPERTY, parse_all, IDR_SV11PROP2IN, IDR_SV11PROP2OUT },
				//SmartViewTestResource{ IDS_STPROPERTY, parse_all, IDR_SV11PROP3IN, IDR_SV11PROP3OUT },
				//SmartViewTestResource{ IDS_STPROPERTY, parse_all, IDR_SV11PROP4IN, IDR_SV11PROP4OUT },
				//SmartViewTestResource{ IDS_STPROPERTY, parse_all, IDR_SV11PROP5IN, IDR_SV11PROP5OUT },
				//SmartViewTestResource{ IDS_STPROPERTY, parse_all, IDR_SV11PROP6IN, IDR_SV11PROP6OUT },
				//SmartViewTestResource{ IDS_STPROPERTY, parse_all, IDR_SV11PROP7IN, IDR_SV11PROP7OUT },
				}));
		}

		TEST_METHOD(Test_STPROPERTYDEFINITIONSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF1IN, IDR_SV12PROPDEF1OUT },
				SmartViewTestResource{ IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF2IN, IDR_SV12PROPDEF2OUT },
				SmartViewTestResource{ IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF3IN, IDR_SV12PROPDEF3OUT },
				SmartViewTestResource{ IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF4IN, IDR_SV12PROPDEF4OUT },
				SmartViewTestResource{ IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF5IN, IDR_SV12PROPDEF5OUT },
				SmartViewTestResource{ IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF6IN, IDR_SV12PROPDEF6OUT },
				}));
		}

		TEST_METHOD(Test_STRECIPIENTROWSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STRECIPIENTROWSTREAM, parse_all, IDR_SV13RECIPROW1IN, IDR_SV13RECIPROW1OUT },
				SmartViewTestResource{ IDS_STRECIPIENTROWSTREAM, parse_all, IDR_SV13RECIPROW2IN, IDR_SV13RECIPROW2OUT },
				}));
		}

		TEST_METHOD(Test_STRECURRENCEPATTERN)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STRECURRENCEPATTERN, parse_all, IDR_SV14ARP1IN, IDR_SV14ARP1OUT },
				SmartViewTestResource{ IDS_STRECURRENCEPATTERN, parse_all, IDR_SV14ARP2IN, IDR_SV14ARP2OUT },
				SmartViewTestResource{ IDS_STRECURRENCEPATTERN, parse_all, IDR_SV14ARP3IN, IDR_SV14ARP3OUT },
				}));
		}

		TEST_METHOD(Test_STREPORTTAG)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STREPORTTAG, parse_all, IDR_SV15REPORTTAG1IN, IDR_SV15REPORTTAG1OUT },
				}));
		}

		TEST_METHOD(Test_STRESTRICTION)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STRESTRICTION, parse_all, IDR_SV16RES1IN, IDR_SV16RES1OUT },
				}));
		}

		TEST_METHOD(Test_STRULECONDITION)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STRULECONDITION, parse_all, IDR_SV17RULECON1IN, IDR_SV17RULECON1OUT },
				SmartViewTestResource{ IDS_STRULECONDITION, parse_all, IDR_SV17RULECON2IN, IDR_SV17RULECON2OUT },
				SmartViewTestResource{ IDS_STRULECONDITION, parse_all, IDR_SV17RULECON3IN, IDR_SV17RULECON3OUT },
				SmartViewTestResource{ IDS_STRULECONDITION, parse_all, IDR_SV17RULECON4IN, IDR_SV17RULECON4OUT },
				SmartViewTestResource{ IDS_STRULECONDITION, parse_all, IDR_SV17RULECON5IN, IDR_SV17RULECON5OUT },
				}));
		}

		TEST_METHOD(Test_STSEARCHFOLDERDEFINITION)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF1IN, IDR_SV18SF1OUT },
				SmartViewTestResource{ IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF2IN, IDR_SV18SF2OUT },
				SmartViewTestResource{ IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF3IN, IDR_SV18SF3OUT },
				SmartViewTestResource{ IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF4IN, IDR_SV18SF4OUT },
				SmartViewTestResource{ IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF5IN, IDR_SV18SF5OUT },
				SmartViewTestResource{ IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF6IN, IDR_SV18SF6OUT },
				}));
		}

		TEST_METHOD(Test_STSECURITYDESCRIPTOR)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STSECURITYDESCRIPTOR, parse_all, IDR_SV19SD1IN, IDR_SV19SD1OUT },
				SmartViewTestResource{ IDS_STSECURITYDESCRIPTOR, parse_all, IDR_SV19SD2IN, IDR_SV19SD2OUT },
				SmartViewTestResource{ IDS_STSECURITYDESCRIPTOR, parse_all, IDR_SV19SD3IN, IDR_SV19SD3OUT },
				SmartViewTestResource{ IDS_STSECURITYDESCRIPTOR, parse_all, IDR_SV19SD4IN, IDR_SV19SD4OUT },
				}));
		}

		TEST_METHOD(Test_STSID)
		{
			test(loadTestData({
				SmartViewTestResource{ IDS_STSID, parse_all, IDR_SV20SID1IN, IDR_SV20SID1OUT },
				SmartViewTestResource{ IDS_STSID, parse_all, IDR_SV20SID2IN, IDR_SV20SID2OUT },
				SmartViewTestResource{ IDS_STSID, parse_all, IDR_SV20SID3IN, IDR_SV20SID3OUT },
				SmartViewTestResource{ IDS_STSID, parse_all, IDR_SV20SID4IN, IDR_SV20SID4OUT },
				SmartViewTestResource{ IDS_STSID, parse_all, IDR_SV20SID5IN, IDR_SV20SID5OUT },
				}));
		}
	};
}