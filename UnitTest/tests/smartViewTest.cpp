#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/smartview/SmartView.h>
#include <core/addin/mfcmapi.h>
#include <core/addin/addin.h>
#include <core/utility/strings.h>

namespace SmartViewTest
{
	TEST_CLASS(SmartViewTest)
	{
	private:
		struct SmartViewTestResource
		{
			parserType structType{};
			bool parseAll{};
			DWORD hex{};
			DWORD expected{};
		};

		struct SmartViewTestData
		{
			parserType structType{};
			bool parseAll{};
			std::wstring testName;
			std::vector<BYTE> hex;
			std::wstring expected;
		};

		void test(std::vector<SmartViewTestData> testData) const
		{
			for (auto& data : testData)
			{
				auto actual = smartview::InterpretBinaryAsString(
					{static_cast<ULONG>(data.hex.size()), data.hex.data()}, data.structType, nullptr);
				unittest::AreEqualEx(data.expected, actual, data.testName.c_str());

				if (data.parseAll)
				{
					for (auto parser : SmartViewParserTypeArray)
					{
						const auto structType = parser.type;
						try
						{
							actual = smartview::InterpretBinaryAsString(
								{static_cast<ULONG>(data.hex.size()), data.hex.data()}, structType, nullptr);
						} catch (const int exception)
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

		static std::vector<SmartViewTestData> loadTestData(
			const std::initializer_list<SmartViewTestResource>& resources)
		{
			static auto handle = GetModuleHandleW(L"UnitTest.dll");
			std::vector<SmartViewTestData> testData;
			for (const auto& resource : resources)
			{
				testData.push_back(SmartViewTestData{resource.structType,
													 resource.parseAll,
													 strings::format(L"%d/%d", resource.hex, resource.expected),
													 strings::HexStringToBin(unittest::loadfile(handle, resource.hex)),
													 unittest::loadfile(handle, resource.expected)});
			}

			return testData;
		}

	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_STADDITIONALRENENTRYIDSEX)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::ADDITIONALRENENTRYIDSEX, unittest::parse_all, IDR_SV1AEI1IN, IDR_SV1AEI1OUT},
				SmartViewTestResource{
					parserType::ADDITIONALRENENTRYIDSEX, unittest::parse_all, IDR_SV1AEI2IN, IDR_SV1AEI2OUT},
				SmartViewTestResource{
					parserType::ADDITIONALRENENTRYIDSEX, unittest::parse_all, IDR_SV1AEI3IN, IDR_SV1AEI3OUT},
				SmartViewTestResource{
					parserType::ADDITIONALRENENTRYIDSEX, unittest::parse_all, IDR_SV1AEI4IN, IDR_SV1AEI4OUT},
			}));
		}

		TEST_METHOD(Test_STAPPOINTMENTRECURRENCEPATTERN)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::APPOINTMENTRECURRENCEPATTERN, unittest::parse_all, IDR_SV2ARP1IN, IDR_SV2ARP1OUT},
				SmartViewTestResource{
					parserType::APPOINTMENTRECURRENCEPATTERN, unittest::parse_all, IDR_SV2ARP2IN, IDR_SV2ARP2OUT},
				SmartViewTestResource{
					parserType::APPOINTMENTRECURRENCEPATTERN, unittest::parse_all, IDR_SV2ARP3IN, IDR_SV2ARP3OUT},
				SmartViewTestResource{
					parserType::APPOINTMENTRECURRENCEPATTERN, unittest::parse_all, IDR_SV2ARP4IN, IDR_SV2ARP4OUT},
			}));
		}

		TEST_METHOD(Test_STCONVERSATIONINDEX)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::CONVERSATIONINDEX, unittest::parse_all, IDR_SV3CI1IN, IDR_SV3CI1OUT},
				SmartViewTestResource{parserType::CONVERSATIONINDEX, unittest::parse_all, IDR_SV3CI2IN, IDR_SV3CI2OUT},
				SmartViewTestResource{parserType::CONVERSATIONINDEX, unittest::parse_all, IDR_SV3CI3IN, IDR_SV3CI3OUT},
				SmartViewTestResource{parserType::CONVERSATIONINDEX, unittest::parse_all, IDR_SV3CI4IN, IDR_SV3CI4OUT},
			}));
		}

		TEST_METHOD(Test_STENTRYID)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID1IN, IDR_SV4EID1OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID2IN, IDR_SV4EID2OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID3IN, IDR_SV4EID3OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID4IN, IDR_SV4EID4OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID5IN, IDR_SV4EID5OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID6IN, IDR_SV4EID6OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID7IN, IDR_SV4EID7OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID8IN, IDR_SV4EID8OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID9IN, IDR_SV4EID9OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID10IN, IDR_SV4EID10OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID11IN, IDR_SV4EID11OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID12IN, IDR_SV4EID12OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID13IN, IDR_SV4EID13OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID14IN, IDR_SV4EID14OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID15IN, IDR_SV4EID15OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID16IN, IDR_SV4EID16OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID17IN, IDR_SV4EID17OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID18IN, IDR_SV4EID18OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID19IN, IDR_SV4EID19OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID20IN, IDR_SV4EID20OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID21IN, IDR_SV4EID21OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID22IN, IDR_SV4EID22OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID23IN, IDR_SV4EID23OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID24IN, IDR_SV4EID24OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID25IN, IDR_SV4EID25OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID26IN, IDR_SV4EID26OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID27IN, IDR_SV4EID27OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID28IN, IDR_SV4EID28OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID29IN, IDR_SV4EID29OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID30IN, IDR_SV4EID30OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID31IN, IDR_SV4EID31OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID32IN, IDR_SV4EID32OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID33IN, IDR_SV4EID33OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID34IN, IDR_SV4EID34OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID35IN, IDR_SV4EID35OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID36IN, IDR_SV4EID36OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID37IN, IDR_SV4EID37OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID38IN, IDR_SV4EID38OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID39IN, IDR_SV4EID39OUT},
				SmartViewTestResource{parserType::ENTRYID, unittest::parse_all, IDR_SV4EID40IN, IDR_SV4EID40OUT},
			}));
		}

		TEST_METHOD(Test_STENTRYLIST)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::ENTRYLIST, unittest::parse_all, IDR_SV5EL1IN, IDR_SV5EL1OUT},
				SmartViewTestResource{parserType::ENTRYLIST, unittest::parse_all, IDR_SV5EL2IN, IDR_SV5EL2OUT},
			}));
		}

		TEST_METHOD(Test_STEXTENDEDFOLDERFLAGS)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::EXTENDEDFOLDERFLAGS, unittest::parse_all, IDR_SV6EFF1IN, IDR_SV6EFF1OUT},
				SmartViewTestResource{
					parserType::EXTENDEDFOLDERFLAGS, unittest::parse_all, IDR_SV6EFF2IN, IDR_SV6EFF2OUT},
			}));
		}

		TEST_METHOD(Test_STEXTENDEDRULECONDITION)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::EXTENDEDRULECONDITION, unittest::parse_all, IDR_SV7EXRULE1IN, IDR_SV7EXRULE1OUT},
				SmartViewTestResource{
					parserType::EXTENDEDRULECONDITION, unittest::parse_all, IDR_SV7EXRULE2IN, IDR_SV7EXRULE2OUT},
				SmartViewTestResource{
					parserType::EXTENDEDRULECONDITION, unittest::parse_all, IDR_SV7EXRULE3IN, IDR_SV7EXRULE3OUT},
				SmartViewTestResource{
					parserType::EXTENDEDRULECONDITION, unittest::parse_all, IDR_SV7EXRULE4IN, IDR_SV7EXRULE4OUT},
			}));
		}

		TEST_METHOD(Test_STFLATENTRYLIST)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::FLATENTRYLIST, unittest::parse_all, IDR_SV8FE1IN, IDR_SV8FE1OUT},
				SmartViewTestResource{parserType::FLATENTRYLIST, unittest::parse_all, IDR_SV8FE2IN, IDR_SV8FE2OUT},
				SmartViewTestResource{parserType::FLATENTRYLIST, unittest::parse_all, IDR_SV8FE3IN, IDR_SV8FE3OUT},
			}));
		}

		TEST_METHOD(Test_STFOLDERUSERFIELDS)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::FOLDERUSERFIELDS, unittest::parse_all, IDR_SV9FUF1IN, IDR_SV9FUF1OUT},
				SmartViewTestResource{parserType::FOLDERUSERFIELDS, unittest::parse_all, IDR_SV9FUF2IN, IDR_SV9FUF2OUT},
				SmartViewTestResource{parserType::FOLDERUSERFIELDS, unittest::parse_all, IDR_SV9FUF3IN, IDR_SV9FUF3OUT},
				SmartViewTestResource{parserType::FOLDERUSERFIELDS, unittest::parse_all, IDR_SV9FUF4IN, IDR_SV9FUF4OUT},
				SmartViewTestResource{parserType::FOLDERUSERFIELDS, unittest::parse_all, IDR_SV9FUF5IN, IDR_SV9FUF5OUT},
			}));
		}

		TEST_METHOD(Test_STGLOBALOBJECTID)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::GLOBALOBJECTID, unittest::parse_all, IDR_SV10GOID1IN, IDR_SV10GOID1OUT},
				SmartViewTestResource{
					parserType::GLOBALOBJECTID, unittest::parse_all, IDR_SV10GOID2IN, IDR_SV10GOID2OUT},
			}));
		}

		TEST_METHOD(Test_STPROPERTY)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::PROPERTIES, unittest::parse_all, IDR_SV11PROP1IN, IDR_SV11PROP1OUT},
				SmartViewTestResource{parserType::PROPERTIES, unittest::parse_all, IDR_SV11PROP2IN, IDR_SV11PROP2OUT},
				SmartViewTestResource{parserType::PROPERTIES, unittest::parse_all, IDR_SV11PROP3IN, IDR_SV11PROP3OUT},
				SmartViewTestResource{parserType::PROPERTIES, unittest::parse_all, IDR_SV11PROP4IN, IDR_SV11PROP4OUT},
				SmartViewTestResource{parserType::PROPERTIES, unittest::parse_all, IDR_SV11PROP5IN, IDR_SV11PROP5OUT},
				SmartViewTestResource{parserType::PROPERTIES, unittest::parse_all, IDR_SV11PROP6IN, IDR_SV11PROP6OUT},
				SmartViewTestResource{parserType::PROPERTIES, unittest::parse_all, IDR_SV11PROP7IN, IDR_SV11PROP7OUT},
				SmartViewTestResource{parserType::PROPERTIES, unittest::parse_all, IDR_SV11PROP8IN, IDR_SV11PROP8OUT},
			}));
		}

		TEST_METHOD(Test_STPROPERTYDEFINITIONSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::PROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF1IN, IDR_SV12PROPDEF1OUT},
				SmartViewTestResource{
					parserType::PROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF2IN, IDR_SV12PROPDEF2OUT},
				SmartViewTestResource{
					parserType::PROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF3IN, IDR_SV12PROPDEF3OUT},
				SmartViewTestResource{
					parserType::PROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF4IN, IDR_SV12PROPDEF4OUT},
				SmartViewTestResource{
					parserType::PROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF5IN, IDR_SV12PROPDEF5OUT},
				SmartViewTestResource{
					parserType::PROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF6IN, IDR_SV12PROPDEF6OUT},
				SmartViewTestResource{
					parserType::PROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF7IN, IDR_SV12PROPDEF7OUT},
				SmartViewTestResource{
					parserType::PROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF8IN, IDR_SV12PROPDEF8OUT},
			}));
		}

		TEST_METHOD(Test_STRECIPIENTROWSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::RECIPIENTROWSTREAM, unittest::parse_all, IDR_SV13RECIPROW1IN, IDR_SV13RECIPROW1OUT},
				SmartViewTestResource{
					parserType::RECIPIENTROWSTREAM, unittest::parse_all, IDR_SV13RECIPROW2IN, IDR_SV13RECIPROW2OUT},
			}));
		}

		TEST_METHOD(Test_STRECURRENCEPATTERN)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::RECURRENCEPATTERN, unittest::parse_all, IDR_SV14ARP1IN, IDR_SV14ARP1OUT},
				SmartViewTestResource{
					parserType::RECURRENCEPATTERN, unittest::parse_all, IDR_SV14ARP2IN, IDR_SV14ARP2OUT},
				SmartViewTestResource{
					parserType::RECURRENCEPATTERN, unittest::parse_all, IDR_SV14ARP3IN, IDR_SV14ARP3OUT},
			}));
		}

		TEST_METHOD(Test_STREPORTTAG)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::REPORTTAG, unittest::parse_all, IDR_SV15REPORTTAG1IN, IDR_SV15REPORTTAG1OUT},
			}));
		}

		TEST_METHOD(Test_STRESTRICTION)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::RESTRICTION, unittest::parse_all, IDR_SV16RES1IN, IDR_SV16RES1OUT},
				SmartViewTestResource{parserType::RESTRICTION, unittest::parse_all, IDR_SV16RES2IN, IDR_SV16RES2OUT},
			}));
		}

		TEST_METHOD(Test_STRULECONDITION)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::RULECONDITION, unittest::parse_all, IDR_SV17RULECON1IN, IDR_SV17RULECON1OUT},
				SmartViewTestResource{
					parserType::RULECONDITION, unittest::parse_all, IDR_SV17RULECON2IN, IDR_SV17RULECON2OUT},
				SmartViewTestResource{
					parserType::RULECONDITION, unittest::parse_all, IDR_SV17RULECON3IN, IDR_SV17RULECON3OUT},
				SmartViewTestResource{
					parserType::RULECONDITION, unittest::parse_all, IDR_SV17RULECON4IN, IDR_SV17RULECON4OUT},
			}));
		}

		TEST_METHOD(Test_STSEARCHFOLDERDEFINITION)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::SEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF1IN, IDR_SV18SF1OUT},
				SmartViewTestResource{
					parserType::SEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF2IN, IDR_SV18SF2OUT},
				SmartViewTestResource{
					parserType::SEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF3IN, IDR_SV18SF3OUT},
				SmartViewTestResource{
					parserType::SEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF4IN, IDR_SV18SF4OUT},
				SmartViewTestResource{
					parserType::SEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF5IN, IDR_SV18SF5OUT},
				SmartViewTestResource{
					parserType::SEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF6IN, IDR_SV18SF6OUT},
				SmartViewTestResource{
					parserType::SEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF7IN, IDR_SV18SF7OUT},
			}));
		}

		TEST_METHOD(Test_STSECURITYDESCRIPTOR)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::SECURITYDESCRIPTOR, unittest::parse_all, IDR_SV19SD1IN, IDR_SV19SD1OUT},
				SmartViewTestResource{
					parserType::SECURITYDESCRIPTOR, unittest::parse_all, IDR_SV19SD2IN, IDR_SV19SD2OUT},
				SmartViewTestResource{
					parserType::SECURITYDESCRIPTOR, unittest::parse_all, IDR_SV19SD3IN, IDR_SV19SD3OUT},
				SmartViewTestResource{
					parserType::SECURITYDESCRIPTOR, unittest::parse_all, IDR_SV19SD4IN, IDR_SV19SD4OUT},
			}));
		}

		TEST_METHOD(Test_STSID)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::SID, unittest::parse_all, IDR_SV20SID1IN, IDR_SV20SID1OUT},
				SmartViewTestResource{parserType::SID, unittest::parse_all, IDR_SV20SID2IN, IDR_SV20SID2OUT},
				SmartViewTestResource{parserType::SID, unittest::parse_all, IDR_SV20SID3IN, IDR_SV20SID3OUT},
				SmartViewTestResource{parserType::SID, unittest::parse_all, IDR_SV20SID4IN, IDR_SV20SID4OUT},
				SmartViewTestResource{parserType::SID, unittest::parse_all, IDR_SV20SID5IN, IDR_SV20SID5OUT},
			}));
		}

		TEST_METHOD(Test_STTASKASSIGNERS)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::TASKASSIGNERS, unittest::parse_all, IDR_SV21TA1IN, IDR_SV21TA1OUT},
			}));
		}

		TEST_METHOD(Test_STTIMEZONE)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::TIMEZONE, unittest::parse_all, IDR_SV22TZ1IN, IDR_SV22TZ1OUT},
			}));
		}

		TEST_METHOD(Test_STTIMEZONEDEFINITION)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::TIMEZONEDEFINITION, unittest::parse_all, IDR_SV23TZD1IN, IDR_SV23TZD1OUT},
				SmartViewTestResource{
					parserType::TIMEZONEDEFINITION, unittest::parse_all, IDR_SV23TZD2IN, IDR_SV23TZD2OUT},
				SmartViewTestResource{
					parserType::TIMEZONEDEFINITION, unittest::parse_all, IDR_SV23TZD3IN, IDR_SV23TZD3OUT},
			}));
		}

		TEST_METHOD(Test_STWEBVIEWPERSISTSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::WEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW1IN, IDR_SV24WEBVIEW1OUT},
				SmartViewTestResource{
					parserType::WEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW2IN, IDR_SV24WEBVIEW2OUT},
				SmartViewTestResource{
					parserType::WEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW3IN, IDR_SV24WEBVIEW3OUT},
				SmartViewTestResource{
					parserType::WEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW4IN, IDR_SV24WEBVIEW4OUT},
				SmartViewTestResource{
					parserType::WEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW5IN, IDR_SV24WEBVIEW5OUT},
				SmartViewTestResource{
					parserType::WEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW6IN, IDR_SV24WEBVIEW6OUT},
				SmartViewTestResource{
					parserType::WEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW7IN, IDR_SV24WEBVIEW7OUT},
			}));
		}

		TEST_METHOD(Test_STNICKNAMECACHE)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::NICKNAMECACHE, unittest::parse_all, IDR_SV25NICKNAME2IN, IDR_SV25NICKNAME2OUT},
				SmartViewTestResource{
					parserType::NICKNAMECACHE, unittest::parse_all, IDR_SV25NICKNAME3IN, IDR_SV25NICKNAME3OUT},
				SmartViewTestResource{
					parserType::NICKNAMECACHE, unittest::parse_all, IDR_SV25NICKNAME4IN, IDR_SV25NICKNAME4OUT},
			}));
		}

		TEST_METHOD(Test_STENCODEENTRYID)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::ENCODEENTRYID, unittest::parse_all, IDR_SV26EIDENCODE1IN, IDR_SV26EIDENCODE1OUT},
			}));
		}

		TEST_METHOD(Test_STDECODEENTRYID)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::DECODEENTRYID, unittest::parse_all, IDR_SV27EIDDECODE1IN, IDR_SV27EIDDECODE1OUT},
			}));
		}

		TEST_METHOD(Test_STVERBSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::VERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM1IN, IDR_SV28VERBSTREAM1OUT},
				SmartViewTestResource{
					parserType::VERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM2IN, IDR_SV28VERBSTREAM2OUT},
				SmartViewTestResource{
					parserType::VERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM3IN, IDR_SV28VERBSTREAM3OUT},
				SmartViewTestResource{
					parserType::VERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM4IN, IDR_SV28VERBSTREAM4OUT},
				SmartViewTestResource{
					parserType::VERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM5IN, IDR_SV28VERBSTREAM5OUT},
				SmartViewTestResource{
					parserType::VERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM6IN, IDR_SV28VERBSTREAM6OUT},
				SmartViewTestResource{
					parserType::VERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM7IN, IDR_SV28VERBSTREAM7OUT},
			}));
		}

		TEST_METHOD(Test_STTOMBSTONE)
		{
			test(loadTestData({
				SmartViewTestResource{
					parserType::TOMBSTONE, unittest::parse_all, IDR_SV29TOMBSTONE1IN, IDR_SV29TOMBSTONE1OUT},
				SmartViewTestResource{
					parserType::TOMBSTONE, unittest::parse_all, IDR_SV29TOMBSTONE2IN, IDR_SV29TOMBSTONE2OUT},
				SmartViewTestResource{
					parserType::TOMBSTONE, unittest::parse_all, IDR_SV29TOMBSTONE3IN, IDR_SV29TOMBSTONE3OUT},
			}));
		}

		TEST_METHOD(Test_STPCL)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::PCL, unittest::parse_all, IDR_SV30PCL1IN, IDR_SV30PCL1OUT},
				SmartViewTestResource{parserType::PCL, unittest::parse_all, IDR_SV30PCL2IN, IDR_SV30PCL2OUT},
				SmartViewTestResource{parserType::PCL, unittest::parse_all, IDR_SV30PCL3IN, IDR_SV30PCL3OUT},
			}));
		}

		TEST_METHOD(Test_STFBSECURITYDESCRIPTOR)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::FBSECURITYDESCRIPTOR,
									  unittest::parse_all,
									  IDR_SV31FREEBUSYSID1IN,
									  IDR_SV31FREEBUSYSID1OUT},
				SmartViewTestResource{parserType::FBSECURITYDESCRIPTOR,
									  unittest::parse_all,
									  IDR_SV31FREEBUSYSID2IN,
									  IDR_SV31FREEBUSYSID2OUT},
			}));
		}

		TEST_METHOD(Test_STXID)
		{
			test(loadTestData({
				SmartViewTestResource{parserType::XID, unittest::parse_all, IDR_SV32XID1IN, IDR_SV32XID1OUT},
				SmartViewTestResource{parserType::XID, unittest::parse_all, IDR_SV32XID2IN, IDR_SV32XID2OUT},
				SmartViewTestResource{parserType::XID, unittest::parse_all, IDR_SV32XID3IN, IDR_SV32XID3OUT},
				SmartViewTestResource{parserType::XID, unittest::parse_all, IDR_SV32XID4IN, IDR_SV32XID4OUT},
			}));
		}
	};
} // namespace SmartViewTest