#include <StdAfx.h>
#include <CppUnitTest.h>
#include <Interpret/SmartView/SmartView.h>
#include <MFCMAPI.h>
#include <UnitTest/resource.h>

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
			std::wstring testName;
			std::vector<BYTE> hex;
			std::wstring expected;
		};

		static const bool parse_all = true;
		static const bool assert_on_failure = true;
		static const bool limit_output = true;
		static const bool ignore_trailing_whitespace = false;

		// Assert::AreEqual doesn't do full logging, so we roll our own
		void AreEqualEx(
			const std::wstring& expected,
			const std::wstring& actual,
			const wchar_t* message = nullptr,
			const __LineInfo* pLineInfo = nullptr) const
		{
			if (ignore_trailing_whitespace)
			{
				if (strings::trimWhitespace(expected) == strings::trimWhitespace(actual)) return;
			}
			else
			{
				if (expected == actual) return;
			}

			if (message != nullptr)
			{
				Logger::WriteMessage(strings::format(L"Test: %ws\n", message).c_str());
			}

			Logger::WriteMessage(L"Diff:\n");

			auto splitExpected = strings::split(expected, L'\n');
			auto splitActual = strings::split(actual, L'\n');
			auto errorCount = 0;
			for (size_t line = 0;
				 line < splitExpected.size() && line < splitActual.size() && (errorCount < 4 || !limit_output);
				 line++)
			{
				if (splitExpected[line] != splitActual[line])
				{
					errorCount++;
					Logger::WriteMessage(strings::format(
											 L"[%d]\n\"%ws\"\n\"%ws\"\n",
											 line + 1,
											 splitExpected[line].c_str(),
											 splitActual[line].c_str())
											 .c_str());
					auto lineErrorCount = 0;
					for (size_t ch = 0; ch < splitExpected[line].size() && ch < splitActual[line].size() &&
										(lineErrorCount < 10 || !limit_output);
						 ch++)
					{
						const auto expectedChar = splitExpected[line][ch];
						const auto actualChar = splitActual[line][ch];
						if (expectedChar != actualChar)
						{
							lineErrorCount++;
							Logger::WriteMessage(strings::format(
													 L"[%d]: %X (%wc) != %X (%wc)\n",
													 ch + 1,
													 expectedChar,
													 expectedChar,
													 actualChar,
													 actualChar)
													 .c_str());
						}
					}
				}
			}

			Logger::WriteMessage(L"\n");
			Logger::WriteMessage(strings::format(L"Expected:\n\"%ws\"\n\n", expected.c_str()).c_str());
			Logger::WriteMessage(strings::format(L"Actual:\n\"%ws\"", actual.c_str()).c_str());

			if (assert_on_failure)
			{
				Assert::Fail(ToString(message).c_str(), pLineInfo);
			}
		}

		void test(std::vector<SmartViewTestData> testData) const
		{
			for (auto data : testData)
			{
				auto actual = smartview::InterpretBinaryAsString(
					{static_cast<ULONG>(data.hex.size()), data.hex.data()}, data.structType, nullptr);
				AreEqualEx(data.expected, actual, data.testName.c_str());

				if (data.parseAll)
				{
					for (ULONG iStruct = IDS_STNOPARSING; iStruct < IDS_STEND; iStruct++)
					{
						const auto structType = static_cast<__ParsingTypeEnum>(iStruct);
						try
						{
							actual = smartview::InterpretBinaryAsString(
								{static_cast<ULONG>(data.hex.size()), data.hex.data()}, structType, nullptr);
						} catch (int exception)
						{
							Logger::WriteMessage(strings::format(
													 L"Testing %ws failed at %ws with error 0x%08X\n",
													 data.testName.c_str(),
													 addin::AddInStructTypeToString(structType).c_str(),
													 exception)
													 .c_str());
							Assert::Fail();
						}
					}
				}
			}
		}

		// Resource files saved in unicode have a byte order mark of 0xfffe
		// We load these in and strip the BOM.
		// Otherwise we load as ansi and convert to unicode
		static std::wstring loadfile(const HMODULE handle, const int name)
		{
			const auto rc = ::FindResource(handle, MAKEINTRESOURCE(name), MAKEINTRESOURCE(TEXTFILE));
			const auto rcData = LoadResource(handle, rc);
			const auto cb = SizeofResource(handle, rc);
			const auto bytes = LockResource(rcData);
			const auto data = static_cast<const BYTE*>(bytes);
			if (cb >= 2 && data[0] == 0xff && data[1] == 0xfe)
			{
				// Skip the byte order mark
				const auto wstr = static_cast<const wchar_t*>(bytes);
				const auto cch = cb / sizeof(wchar_t);
				return std::wstring(wstr + 1, cch - 1);
			}

			const auto str = std::string(static_cast<const char*>(bytes), cb);
			return strings::stringTowstring(str);
		}

		static std::vector<SmartViewTestData> loadTestData(std::initializer_list<SmartViewTestResource> resources)
		{
			static auto handle = GetModuleHandleW(L"UnitTest.dll");
			std::vector<SmartViewTestData> testData;
			for (auto resource : resources)
			{
				testData.push_back(SmartViewTestData{resource.structType,
													 resource.parseAll,
													 strings::format(L"%d/%d", resource.hex, resource.expected),
													 strings::HexStringToBin(loadfile(handle, resource.hex)),
													 loadfile(handle, resource.expected)});
			}

			return testData;
		}

	public:
		TEST_CLASS_INITIALIZE(Initialize_smartview)
		{
			// Set up our property arrays or nothing works
			addin::MergeAddInArrays();

			registry::RegKeys[registry::regkeyDO_SMART_VIEW].ulCurDWORD = 1;
			registry::RegKeys[registry::regkeyUSE_GETPROPLIST].ulCurDWORD = 1;
			registry::RegKeys[registry::regkeyPARSED_NAMED_PROPS].ulCurDWORD = 1;
			registry::RegKeys[registry::regkeyCACHE_NAME_DPROPS].ulCurDWORD = 1;

			strings::setTestInstance(GetModuleHandleW(L"UnitTest.dll"));
		}

		TEST_METHOD(Test_STADDITIONALRENENTRYIDSEX)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STADDITIONALRENENTRYIDSEX, parse_all, IDR_SV1AEI1IN, IDR_SV1AEI1OUT},
				SmartViewTestResource{IDS_STADDITIONALRENENTRYIDSEX, parse_all, IDR_SV1AEI2IN, IDR_SV1AEI2OUT},
			}));
		}

		TEST_METHOD(Test_STAPPOINTMENTRECURRENCEPATTERN)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STAPPOINTMENTRECURRENCEPATTERN, parse_all, IDR_SV2ARP1IN, IDR_SV2ARP1OUT},
				SmartViewTestResource{IDS_STAPPOINTMENTRECURRENCEPATTERN, parse_all, IDR_SV2ARP2IN, IDR_SV2ARP2OUT},
				SmartViewTestResource{IDS_STAPPOINTMENTRECURRENCEPATTERN, parse_all, IDR_SV2ARP3IN, IDR_SV2ARP3OUT},
				SmartViewTestResource{IDS_STAPPOINTMENTRECURRENCEPATTERN, parse_all, IDR_SV2ARP4IN, IDR_SV2ARP4OUT},
			}));
		}

		TEST_METHOD(Test_STCONVERSATIONINDEX)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STCONVERSATIONINDEX, parse_all, IDR_SV3CI1IN, IDR_SV3CI1OUT},
				SmartViewTestResource{IDS_STCONVERSATIONINDEX, parse_all, IDR_SV3CI2IN, IDR_SV3CI2OUT},
				SmartViewTestResource{IDS_STCONVERSATIONINDEX, parse_all, IDR_SV3CI3IN, IDR_SV3CI3OUT},
			}));
		}

		TEST_METHOD(Test_STENTRYID)
		{
			test(loadTestData({SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID1IN, IDR_SV4EID1OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID2IN, IDR_SV4EID2OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID3IN, IDR_SV4EID3OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID4IN, IDR_SV4EID4OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID5IN, IDR_SV4EID5OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID6IN, IDR_SV4EID6OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID7IN, IDR_SV4EID7OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID8IN, IDR_SV4EID8OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID9IN, IDR_SV4EID9OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID10IN, IDR_SV4EID10OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID11IN, IDR_SV4EID11OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID12IN, IDR_SV4EID12OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID13IN, IDR_SV4EID13OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID14IN, IDR_SV4EID14OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID15IN, IDR_SV4EID15OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID16IN, IDR_SV4EID16OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID17IN, IDR_SV4EID17OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID18IN, IDR_SV4EID18OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID19IN, IDR_SV4EID19OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID20IN, IDR_SV4EID20OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID21IN, IDR_SV4EID21OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID22IN, IDR_SV4EID22OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID23IN, IDR_SV4EID23OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID24IN, IDR_SV4EID24OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID25IN, IDR_SV4EID25OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID26IN, IDR_SV4EID26OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID27IN, IDR_SV4EID27OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID28IN, IDR_SV4EID28OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID29IN, IDR_SV4EID29OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID30IN, IDR_SV4EID30OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID31IN, IDR_SV4EID31OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID32IN, IDR_SV4EID32OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID33IN, IDR_SV4EID33OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID34IN, IDR_SV4EID34OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID35IN, IDR_SV4EID35OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID36IN, IDR_SV4EID36OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID37IN, IDR_SV4EID37OUT},
							   SmartViewTestResource{IDS_STENTRYID, parse_all, IDR_SV4EID38IN, IDR_SV4EID38OUT}}));
		}

		TEST_METHOD(Test_STENTRYLIST)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STENTRYLIST, parse_all, IDR_SV5EL1IN, IDR_SV5EL1OUT},
				SmartViewTestResource{IDS_STENTRYLIST, parse_all, IDR_SV5EL2IN, IDR_SV5EL2OUT},
			}));
		}

		TEST_METHOD(Test_STEXTENDEDFOLDERFLAGS)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STEXTENDEDFOLDERFLAGS, parse_all, IDR_SV6EFF1IN, IDR_SV6EFF1OUT},
				SmartViewTestResource{IDS_STEXTENDEDFOLDERFLAGS, parse_all, IDR_SV6EFF2IN, IDR_SV6EFF2OUT},
			}));
		}

		TEST_METHOD(Test_STEXTENDEDRULECONDITION)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STEXTENDEDRULECONDITION, parse_all, IDR_SV7EXRULE1IN, IDR_SV7EXRULE1OUT},
				SmartViewTestResource{IDS_STEXTENDEDRULECONDITION, parse_all, IDR_SV7EXRULE2IN, IDR_SV7EXRULE2OUT},
				SmartViewTestResource{IDS_STEXTENDEDRULECONDITION, parse_all, IDR_SV7EXRULE3IN, IDR_SV7EXRULE3OUT},
				SmartViewTestResource{IDS_STEXTENDEDRULECONDITION, parse_all, IDR_SV7EXRULE4IN, IDR_SV7EXRULE4OUT},
			}));
		}

		TEST_METHOD(Test_STFLATENTRYLIST)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STFLATENTRYLIST, parse_all, IDR_SV8FE1IN, IDR_SV8FE1OUT},
				SmartViewTestResource{IDS_STFLATENTRYLIST, parse_all, IDR_SV8FE2IN, IDR_SV8FE2OUT},
			}));
		}

		TEST_METHOD(Test_STFOLDERUSERFIELDS)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STFOLDERUSERFIELDS, parse_all, IDR_SV9FUF1IN, IDR_SV9FUF1OUT},
				SmartViewTestResource{IDS_STFOLDERUSERFIELDS, parse_all, IDR_SV9FUF2IN, IDR_SV9FUF2OUT},
			}));
		}

		TEST_METHOD(Test_STGLOBALOBJECTID)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STGLOBALOBJECTID, parse_all, IDR_SV10GOID1IN, IDR_SV10GOID1OUT},
				SmartViewTestResource{IDS_STGLOBALOBJECTID, parse_all, IDR_SV10GOID2IN, IDR_SV10GOID2OUT},
			}));
		}

		TEST_METHOD(Test_STPROPERTY)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STPROPERTIES, parse_all, IDR_SV11PROP1IN, IDR_SV11PROP1OUT},
				SmartViewTestResource{IDS_STPROPERTIES, parse_all, IDR_SV11PROP2IN, IDR_SV11PROP2OUT},
				SmartViewTestResource{IDS_STPROPERTIES, parse_all, IDR_SV11PROP3IN, IDR_SV11PROP3OUT},
				SmartViewTestResource{IDS_STPROPERTIES, parse_all, IDR_SV11PROP4IN, IDR_SV11PROP4OUT},
				SmartViewTestResource{IDS_STPROPERTIES, parse_all, IDR_SV11PROP5IN, IDR_SV11PROP5OUT},
				SmartViewTestResource{IDS_STPROPERTIES, parse_all, IDR_SV11PROP6IN, IDR_SV11PROP6OUT},
				SmartViewTestResource{IDS_STPROPERTIES, parse_all, IDR_SV11PROP7IN, IDR_SV11PROP7OUT},
			}));
		}

		TEST_METHOD(Test_STPROPERTYDEFINITIONSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF1IN, IDR_SV12PROPDEF1OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF2IN, IDR_SV12PROPDEF2OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF3IN, IDR_SV12PROPDEF3OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF4IN, IDR_SV12PROPDEF4OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF5IN, IDR_SV12PROPDEF5OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, parse_all, IDR_SV12PROPDEF6IN, IDR_SV12PROPDEF6OUT},
			}));
		}

		TEST_METHOD(Test_STRECIPIENTROWSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STRECIPIENTROWSTREAM, parse_all, IDR_SV13RECIPROW1IN, IDR_SV13RECIPROW1OUT},
				SmartViewTestResource{IDS_STRECIPIENTROWSTREAM, parse_all, IDR_SV13RECIPROW2IN, IDR_SV13RECIPROW2OUT},
			}));
		}

		TEST_METHOD(Test_STRECURRENCEPATTERN)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STRECURRENCEPATTERN, parse_all, IDR_SV14ARP1IN, IDR_SV14ARP1OUT},
				SmartViewTestResource{IDS_STRECURRENCEPATTERN, parse_all, IDR_SV14ARP2IN, IDR_SV14ARP2OUT},
				SmartViewTestResource{IDS_STRECURRENCEPATTERN, parse_all, IDR_SV14ARP3IN, IDR_SV14ARP3OUT},
			}));
		}

		TEST_METHOD(Test_STREPORTTAG)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STREPORTTAG, parse_all, IDR_SV15REPORTTAG1IN, IDR_SV15REPORTTAG1OUT},
			}));
		}

		TEST_METHOD(Test_STRESTRICTION)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STRESTRICTION, parse_all, IDR_SV16RES1IN, IDR_SV16RES1OUT},
			}));
		}

		TEST_METHOD(Test_STRULECONDITION)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STRULECONDITION, parse_all, IDR_SV17RULECON1IN, IDR_SV17RULECON1OUT},
				SmartViewTestResource{IDS_STRULECONDITION, parse_all, IDR_SV17RULECON2IN, IDR_SV17RULECON2OUT},
				SmartViewTestResource{IDS_STRULECONDITION, parse_all, IDR_SV17RULECON3IN, IDR_SV17RULECON3OUT},
				SmartViewTestResource{IDS_STRULECONDITION, parse_all, IDR_SV17RULECON4IN, IDR_SV17RULECON4OUT},
				SmartViewTestResource{IDS_STRULECONDITION, parse_all, IDR_SV17RULECON5IN, IDR_SV17RULECON5OUT},
			}));
		}

		TEST_METHOD(Test_STSEARCHFOLDERDEFINITION)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF1IN, IDR_SV18SF1OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF2IN, IDR_SV18SF2OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF3IN, IDR_SV18SF3OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF4IN, IDR_SV18SF4OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF5IN, IDR_SV18SF5OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, parse_all, IDR_SV18SF6IN, IDR_SV18SF6OUT},
			}));
		}

		TEST_METHOD(Test_STSECURITYDESCRIPTOR)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STSECURITYDESCRIPTOR, parse_all, IDR_SV19SD1IN, IDR_SV19SD1OUT},
				SmartViewTestResource{IDS_STSECURITYDESCRIPTOR, parse_all, IDR_SV19SD2IN, IDR_SV19SD2OUT},
				SmartViewTestResource{IDS_STSECURITYDESCRIPTOR, parse_all, IDR_SV19SD3IN, IDR_SV19SD3OUT},
				SmartViewTestResource{IDS_STSECURITYDESCRIPTOR, parse_all, IDR_SV19SD4IN, IDR_SV19SD4OUT},
			}));
		}

		TEST_METHOD(Test_STSID)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STSID, parse_all, IDR_SV20SID1IN, IDR_SV20SID1OUT},
				SmartViewTestResource{IDS_STSID, parse_all, IDR_SV20SID2IN, IDR_SV20SID2OUT},
				SmartViewTestResource{IDS_STSID, parse_all, IDR_SV20SID3IN, IDR_SV20SID3OUT},
				SmartViewTestResource{IDS_STSID, parse_all, IDR_SV20SID4IN, IDR_SV20SID4OUT},
				SmartViewTestResource{IDS_STSID, parse_all, IDR_SV20SID5IN, IDR_SV20SID5OUT},
			}));
		}

		TEST_METHOD(Test_STTASKASSIGNERS)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STTASKASSIGNERS, parse_all, IDR_SV21TA1IN, IDR_SV21TA1OUT},
			}));
		}

		TEST_METHOD(Test_STTIMEZONE)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STTIMEZONE, parse_all, IDR_SV22TZ1IN, IDR_SV22TZ1OUT},
			}));
		}

		TEST_METHOD(Test_STTIMEZONEDEFINITION)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STTIMEZONEDEFINITION, parse_all, IDR_SV23TZD1IN, IDR_SV23TZD1OUT},
				SmartViewTestResource{IDS_STTIMEZONEDEFINITION, parse_all, IDR_SV23TZD2IN, IDR_SV23TZD2OUT},
			}));
		}

		TEST_METHOD(Test_STWEBVIEWPERSISTSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STWEBVIEWPERSISTSTREAM, parse_all, IDR_SV24WEBVIEW1IN, IDR_SV24WEBVIEW1OUT},
				SmartViewTestResource{IDS_STWEBVIEWPERSISTSTREAM, parse_all, IDR_SV24WEBVIEW2IN, IDR_SV24WEBVIEW2OUT},
				SmartViewTestResource{IDS_STWEBVIEWPERSISTSTREAM, parse_all, IDR_SV24WEBVIEW3IN, IDR_SV24WEBVIEW3OUT},
				SmartViewTestResource{IDS_STWEBVIEWPERSISTSTREAM, parse_all, IDR_SV24WEBVIEW4IN, IDR_SV24WEBVIEW4OUT},
				SmartViewTestResource{IDS_STWEBVIEWPERSISTSTREAM, parse_all, IDR_SV24WEBVIEW5IN, IDR_SV24WEBVIEW5OUT},
				SmartViewTestResource{IDS_STWEBVIEWPERSISTSTREAM, parse_all, IDR_SV24WEBVIEW6IN, IDR_SV24WEBVIEW6OUT},
			}));
		}

		TEST_METHOD(Test_STNICKNAMECACHE)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STNICKNAMECACHE, parse_all, IDR_SV25NICKNAME2IN, IDR_SV25NICKNAME2OUT},
			}));
		}

		TEST_METHOD(Test_STENCODEENTRYID)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STENCODEENTRYID, parse_all, IDR_SV26EIDENCODE1IN, IDR_SV26EIDENCODE1OUT},
			}));
		}

		TEST_METHOD(Test_STDECODEENTRYID)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STDECODEENTRYID, parse_all, IDR_SV27EIDDECODE1IN, IDR_SV27EIDDECODE1OUT},
			}));
		}

		TEST_METHOD(Test_STVERBSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STVERBSTREAM, parse_all, IDR_SV28VERBSTREAM1IN, IDR_SV28VERBSTREAM1OUT},
				SmartViewTestResource{IDS_STVERBSTREAM, parse_all, IDR_SV28VERBSTREAM2IN, IDR_SV28VERBSTREAM2OUT},
				SmartViewTestResource{IDS_STVERBSTREAM, parse_all, IDR_SV28VERBSTREAM3IN, IDR_SV28VERBSTREAM3OUT},
				SmartViewTestResource{IDS_STVERBSTREAM, parse_all, IDR_SV28VERBSTREAM4IN, IDR_SV28VERBSTREAM4OUT},
				SmartViewTestResource{IDS_STVERBSTREAM, parse_all, IDR_SV28VERBSTREAM5IN, IDR_SV28VERBSTREAM5OUT},
				SmartViewTestResource{IDS_STVERBSTREAM, parse_all, IDR_SV28VERBSTREAM6IN, IDR_SV28VERBSTREAM6OUT},
			}));
		}

		TEST_METHOD(Test_STTOMBSTONE)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STTOMBSTONE, parse_all, IDR_SV29TOMBSTONE1IN, IDR_SV29TOMBSTONE1OUT},
				SmartViewTestResource{IDS_STTOMBSTONE, parse_all, IDR_SV29TOMBSTONE2IN, IDR_SV29TOMBSTONE2OUT},
				SmartViewTestResource{IDS_STTOMBSTONE, parse_all, IDR_SV29TOMBSTONE3IN, IDR_SV29TOMBSTONE3OUT},
			}));
		}

		TEST_METHOD(Test_STPCL)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STPCL, parse_all, IDR_SV30PCL1IN, IDR_SV30PCL1OUT},
				SmartViewTestResource{IDS_STPCL, parse_all, IDR_SV30PCL2IN, IDR_SV30PCL2OUT},
				SmartViewTestResource{IDS_STPCL, parse_all, IDR_SV30PCL3IN, IDR_SV30PCL3OUT},
			}));
		}

		TEST_METHOD(Test_STFBSECURITYDESCRIPTOR)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STFBSECURITYDESCRIPTOR, parse_all, IDR_SV31FREEBUSYSID1IN, IDR_SV31FREEBUSYSID1OUT},
				SmartViewTestResource{
					IDS_STFBSECURITYDESCRIPTOR, parse_all, IDR_SV31FREEBUSYSID2IN, IDR_SV31FREEBUSYSID2OUT},
			}));
		}

		TEST_METHOD(Test_STXID)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STXID, parse_all, IDR_SV32XID1IN, IDR_SV32XID1OUT},
				SmartViewTestResource{IDS_STXID, parse_all, IDR_SV32XID2IN, IDR_SV32XID2OUT},
				SmartViewTestResource{IDS_STXID, parse_all, IDR_SV32XID3IN, IDR_SV32XID3OUT},
				SmartViewTestResource{IDS_STXID, parse_all, IDR_SV32XID4IN, IDR_SV32XID4OUT},
			}));
		}
	};
} // namespace SmartViewTest