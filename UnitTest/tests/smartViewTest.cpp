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
			__ParsingTypeEnum structType{};
			bool parseAll{};
			DWORD hex{};
			DWORD expected{};
		};

		struct SmartViewTestData
		{
			__ParsingTypeEnum structType{};
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
					IDS_STADDITIONALRENENTRYIDSEX, unittest::parse_all, IDR_SV1AEI1IN, IDR_SV1AEI1OUT},
				SmartViewTestResource{
					IDS_STADDITIONALRENENTRYIDSEX, unittest::parse_all, IDR_SV1AEI2IN, IDR_SV1AEI2OUT},
			}));
		}

		TEST_METHOD(Test_STAPPOINTMENTRECURRENCEPATTERN)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STAPPOINTMENTRECURRENCEPATTERN, unittest::parse_all, IDR_SV2ARP1IN, IDR_SV2ARP1OUT},
				SmartViewTestResource{
					IDS_STAPPOINTMENTRECURRENCEPATTERN, unittest::parse_all, IDR_SV2ARP2IN, IDR_SV2ARP2OUT},
				SmartViewTestResource{
					IDS_STAPPOINTMENTRECURRENCEPATTERN, unittest::parse_all, IDR_SV2ARP3IN, IDR_SV2ARP3OUT},
				SmartViewTestResource{
					IDS_STAPPOINTMENTRECURRENCEPATTERN, unittest::parse_all, IDR_SV2ARP4IN, IDR_SV2ARP4OUT},
			}));
		}

		TEST_METHOD(Test_STCONVERSATIONINDEX)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STCONVERSATIONINDEX, unittest::parse_all, IDR_SV3CI1IN, IDR_SV3CI1OUT},
				SmartViewTestResource{IDS_STCONVERSATIONINDEX, unittest::parse_all, IDR_SV3CI2IN, IDR_SV3CI2OUT},
				SmartViewTestResource{IDS_STCONVERSATIONINDEX, unittest::parse_all, IDR_SV3CI3IN, IDR_SV3CI3OUT},
			}));
		}

		TEST_METHOD(Test_STENTRYID)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID1IN, IDR_SV4EID1OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID2IN, IDR_SV4EID2OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID3IN, IDR_SV4EID3OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID4IN, IDR_SV4EID4OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID5IN, IDR_SV4EID5OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID6IN, IDR_SV4EID6OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID7IN, IDR_SV4EID7OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID8IN, IDR_SV4EID8OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID9IN, IDR_SV4EID9OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID10IN, IDR_SV4EID10OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID11IN, IDR_SV4EID11OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID12IN, IDR_SV4EID12OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID13IN, IDR_SV4EID13OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID14IN, IDR_SV4EID14OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID15IN, IDR_SV4EID15OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID16IN, IDR_SV4EID16OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID17IN, IDR_SV4EID17OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID18IN, IDR_SV4EID18OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID19IN, IDR_SV4EID19OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID20IN, IDR_SV4EID20OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID21IN, IDR_SV4EID21OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID22IN, IDR_SV4EID22OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID23IN, IDR_SV4EID23OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID24IN, IDR_SV4EID24OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID25IN, IDR_SV4EID25OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID26IN, IDR_SV4EID26OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID27IN, IDR_SV4EID27OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID28IN, IDR_SV4EID28OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID29IN, IDR_SV4EID29OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID30IN, IDR_SV4EID30OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID31IN, IDR_SV4EID31OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID32IN, IDR_SV4EID32OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID33IN, IDR_SV4EID33OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID34IN, IDR_SV4EID34OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID35IN, IDR_SV4EID35OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID36IN, IDR_SV4EID36OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID37IN, IDR_SV4EID37OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID38IN, IDR_SV4EID38OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID39IN, IDR_SV4EID39OUT},
				SmartViewTestResource{IDS_STENTRYID, unittest::parse_all, IDR_SV4EID40IN, IDR_SV4EID40OUT},
			}));
		}

		TEST_METHOD(Test_STENTRYLIST)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STENTRYLIST, unittest::parse_all, IDR_SV5EL1IN, IDR_SV5EL1OUT},
				SmartViewTestResource{IDS_STENTRYLIST, unittest::parse_all, IDR_SV5EL2IN, IDR_SV5EL2OUT},
			}));
		}

		TEST_METHOD(Test_STEXTENDEDFOLDERFLAGS)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STEXTENDEDFOLDERFLAGS, unittest::parse_all, IDR_SV6EFF1IN, IDR_SV6EFF1OUT},
				SmartViewTestResource{IDS_STEXTENDEDFOLDERFLAGS, unittest::parse_all, IDR_SV6EFF2IN, IDR_SV6EFF2OUT},
			}));
		}

		TEST_METHOD(Test_STEXTENDEDRULECONDITION)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STEXTENDEDRULECONDITION, unittest::parse_all, IDR_SV7EXRULE1IN, IDR_SV7EXRULE1OUT},
				SmartViewTestResource{
					IDS_STEXTENDEDRULECONDITION, unittest::parse_all, IDR_SV7EXRULE2IN, IDR_SV7EXRULE2OUT},
				SmartViewTestResource{
					IDS_STEXTENDEDRULECONDITION, unittest::parse_all, IDR_SV7EXRULE3IN, IDR_SV7EXRULE3OUT},
				SmartViewTestResource{
					IDS_STEXTENDEDRULECONDITION, unittest::parse_all, IDR_SV7EXRULE4IN, IDR_SV7EXRULE4OUT},
			}));
		}

		TEST_METHOD(Test_STFLATENTRYLIST)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STFLATENTRYLIST, unittest::parse_all, IDR_SV8FE1IN, IDR_SV8FE1OUT},
				SmartViewTestResource{IDS_STFLATENTRYLIST, unittest::parse_all, IDR_SV8FE2IN, IDR_SV8FE2OUT},
			}));
		}

		TEST_METHOD(Test_STFOLDERUSERFIELDS)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STFOLDERUSERFIELDS, unittest::parse_all, IDR_SV9FUF1IN, IDR_SV9FUF1OUT},
				SmartViewTestResource{IDS_STFOLDERUSERFIELDS, unittest::parse_all, IDR_SV9FUF2IN, IDR_SV9FUF2OUT},
			}));
		}

		TEST_METHOD(Test_STGLOBALOBJECTID)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STGLOBALOBJECTID, unittest::parse_all, IDR_SV10GOID1IN, IDR_SV10GOID1OUT},
				SmartViewTestResource{IDS_STGLOBALOBJECTID, unittest::parse_all, IDR_SV10GOID2IN, IDR_SV10GOID2OUT},
			}));
		}

		TEST_METHOD(Test_STPROPERTY)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STPROPERTIES, unittest::parse_all, IDR_SV11PROP1IN, IDR_SV11PROP1OUT},
				SmartViewTestResource{IDS_STPROPERTIES, unittest::parse_all, IDR_SV11PROP2IN, IDR_SV11PROP2OUT},
				SmartViewTestResource{IDS_STPROPERTIES, unittest::parse_all, IDR_SV11PROP3IN, IDR_SV11PROP3OUT},
				SmartViewTestResource{IDS_STPROPERTIES, unittest::parse_all, IDR_SV11PROP4IN, IDR_SV11PROP4OUT},
				SmartViewTestResource{IDS_STPROPERTIES, unittest::parse_all, IDR_SV11PROP5IN, IDR_SV11PROP5OUT},
				SmartViewTestResource{IDS_STPROPERTIES, unittest::parse_all, IDR_SV11PROP6IN, IDR_SV11PROP6OUT},
				SmartViewTestResource{IDS_STPROPERTIES, unittest::parse_all, IDR_SV11PROP7IN, IDR_SV11PROP7OUT},
			}));
		}

		TEST_METHOD(Test_STPROPERTYDEFINITIONSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF1IN, IDR_SV12PROPDEF1OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF2IN, IDR_SV12PROPDEF2OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF3IN, IDR_SV12PROPDEF3OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF4IN, IDR_SV12PROPDEF4OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF5IN, IDR_SV12PROPDEF5OUT},
				SmartViewTestResource{
					IDS_STPROPERTYDEFINITIONSTREAM, unittest::parse_all, IDR_SV12PROPDEF6IN, IDR_SV12PROPDEF6OUT},
			}));
		}

		TEST_METHOD(Test_STRECIPIENTROWSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STRECIPIENTROWSTREAM, unittest::parse_all, IDR_SV13RECIPROW1IN, IDR_SV13RECIPROW1OUT},
				SmartViewTestResource{
					IDS_STRECIPIENTROWSTREAM, unittest::parse_all, IDR_SV13RECIPROW2IN, IDR_SV13RECIPROW2OUT},
			}));
		}

		TEST_METHOD(Test_STRECURRENCEPATTERN)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STRECURRENCEPATTERN, unittest::parse_all, IDR_SV14ARP1IN, IDR_SV14ARP1OUT},
				SmartViewTestResource{IDS_STRECURRENCEPATTERN, unittest::parse_all, IDR_SV14ARP2IN, IDR_SV14ARP2OUT},
				SmartViewTestResource{IDS_STRECURRENCEPATTERN, unittest::parse_all, IDR_SV14ARP3IN, IDR_SV14ARP3OUT},
			}));
		}

		TEST_METHOD(Test_STREPORTTAG)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STREPORTTAG, unittest::parse_all, IDR_SV15REPORTTAG1IN, IDR_SV15REPORTTAG1OUT},
			}));
		}

		TEST_METHOD(Test_STRESTRICTION)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STRESTRICTION, unittest::parse_all, IDR_SV16RES1IN, IDR_SV16RES1OUT},
			}));
		}

		TEST_METHOD(Test_STRULECONDITION)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STRULECONDITION, unittest::parse_all, IDR_SV17RULECON1IN, IDR_SV17RULECON1OUT},
				SmartViewTestResource{
					IDS_STRULECONDITION, unittest::parse_all, IDR_SV17RULECON2IN, IDR_SV17RULECON2OUT},
				SmartViewTestResource{
					IDS_STRULECONDITION, unittest::parse_all, IDR_SV17RULECON3IN, IDR_SV17RULECON3OUT},
				SmartViewTestResource{
					IDS_STRULECONDITION, unittest::parse_all, IDR_SV17RULECON4IN, IDR_SV17RULECON4OUT},
				SmartViewTestResource{
					IDS_STRULECONDITION, unittest::parse_all, IDR_SV17RULECON5IN, IDR_SV17RULECON5OUT},
			}));
		}

		TEST_METHOD(Test_STSEARCHFOLDERDEFINITION)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF1IN, IDR_SV18SF1OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF2IN, IDR_SV18SF2OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF3IN, IDR_SV18SF3OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF4IN, IDR_SV18SF4OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF5IN, IDR_SV18SF5OUT},
				SmartViewTestResource{IDS_STSEARCHFOLDERDEFINITION, unittest::parse_all, IDR_SV18SF6IN, IDR_SV18SF6OUT},
			}));
		}

		TEST_METHOD(Test_STSECURITYDESCRIPTOR)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STSECURITYDESCRIPTOR, unittest::parse_all, IDR_SV19SD1IN, IDR_SV19SD1OUT},
				SmartViewTestResource{IDS_STSECURITYDESCRIPTOR, unittest::parse_all, IDR_SV19SD2IN, IDR_SV19SD2OUT},
				SmartViewTestResource{IDS_STSECURITYDESCRIPTOR, unittest::parse_all, IDR_SV19SD3IN, IDR_SV19SD3OUT},
				SmartViewTestResource{IDS_STSECURITYDESCRIPTOR, unittest::parse_all, IDR_SV19SD4IN, IDR_SV19SD4OUT},
			}));
		}

		TEST_METHOD(Test_STSID)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STSID, unittest::parse_all, IDR_SV20SID1IN, IDR_SV20SID1OUT},
				SmartViewTestResource{IDS_STSID, unittest::parse_all, IDR_SV20SID2IN, IDR_SV20SID2OUT},
				SmartViewTestResource{IDS_STSID, unittest::parse_all, IDR_SV20SID3IN, IDR_SV20SID3OUT},
				SmartViewTestResource{IDS_STSID, unittest::parse_all, IDR_SV20SID4IN, IDR_SV20SID4OUT},
				SmartViewTestResource{IDS_STSID, unittest::parse_all, IDR_SV20SID5IN, IDR_SV20SID5OUT},
			}));
		}

		TEST_METHOD(Test_STTASKASSIGNERS)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STTASKASSIGNERS, unittest::parse_all, IDR_SV21TA1IN, IDR_SV21TA1OUT},
			}));
		}

		TEST_METHOD(Test_STTIMEZONE)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STTIMEZONE, unittest::parse_all, IDR_SV22TZ1IN, IDR_SV22TZ1OUT},
			}));
		}

		TEST_METHOD(Test_STTIMEZONEDEFINITION)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STTIMEZONEDEFINITION, unittest::parse_all, IDR_SV23TZD1IN, IDR_SV23TZD1OUT},
				SmartViewTestResource{IDS_STTIMEZONEDEFINITION, unittest::parse_all, IDR_SV23TZD2IN, IDR_SV23TZD2OUT},
				SmartViewTestResource{IDS_STTIMEZONEDEFINITION, unittest::parse_all, IDR_SV23TZD3IN, IDR_SV23TZD3OUT},
			}));
		}

		TEST_METHOD(Test_STWEBVIEWPERSISTSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STWEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW1IN, IDR_SV24WEBVIEW1OUT},
				SmartViewTestResource{
					IDS_STWEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW2IN, IDR_SV24WEBVIEW2OUT},
				SmartViewTestResource{
					IDS_STWEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW3IN, IDR_SV24WEBVIEW3OUT},
				SmartViewTestResource{
					IDS_STWEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW4IN, IDR_SV24WEBVIEW4OUT},
				SmartViewTestResource{
					IDS_STWEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW5IN, IDR_SV24WEBVIEW5OUT},
				SmartViewTestResource{
					IDS_STWEBVIEWPERSISTSTREAM, unittest::parse_all, IDR_SV24WEBVIEW6IN, IDR_SV24WEBVIEW6OUT},
			}));
		}

		TEST_METHOD(Test_STNICKNAMECACHE)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STNICKNAMECACHE, unittest::parse_all, IDR_SV25NICKNAME2IN, IDR_SV25NICKNAME2OUT},
			}));
		}

		TEST_METHOD(Test_STENCODEENTRYID)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STENCODEENTRYID, unittest::parse_all, IDR_SV26EIDENCODE1IN, IDR_SV26EIDENCODE1OUT},
			}));
		}

		TEST_METHOD(Test_STDECODEENTRYID)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STDECODEENTRYID, unittest::parse_all, IDR_SV27EIDDECODE1IN, IDR_SV27EIDDECODE1OUT},
			}));
		}

		TEST_METHOD(Test_STVERBSTREAM)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STVERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM1IN, IDR_SV28VERBSTREAM1OUT},
				SmartViewTestResource{
					IDS_STVERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM2IN, IDR_SV28VERBSTREAM2OUT},
				SmartViewTestResource{
					IDS_STVERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM3IN, IDR_SV28VERBSTREAM3OUT},
				SmartViewTestResource{
					IDS_STVERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM4IN, IDR_SV28VERBSTREAM4OUT},
				SmartViewTestResource{
					IDS_STVERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM5IN, IDR_SV28VERBSTREAM5OUT},
				SmartViewTestResource{
					IDS_STVERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM6IN, IDR_SV28VERBSTREAM6OUT},
				SmartViewTestResource{
					IDS_STVERBSTREAM, unittest::parse_all, IDR_SV28VERBSTREAM7IN, IDR_SV28VERBSTREAM7OUT},
			}));
		}

		TEST_METHOD(Test_STTOMBSTONE)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STTOMBSTONE, unittest::parse_all, IDR_SV29TOMBSTONE1IN, IDR_SV29TOMBSTONE1OUT},
				SmartViewTestResource{
					IDS_STTOMBSTONE, unittest::parse_all, IDR_SV29TOMBSTONE2IN, IDR_SV29TOMBSTONE2OUT},
				SmartViewTestResource{
					IDS_STTOMBSTONE, unittest::parse_all, IDR_SV29TOMBSTONE3IN, IDR_SV29TOMBSTONE3OUT},
			}));
		}

		TEST_METHOD(Test_STPCL)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STPCL, unittest::parse_all, IDR_SV30PCL1IN, IDR_SV30PCL1OUT},
				SmartViewTestResource{IDS_STPCL, unittest::parse_all, IDR_SV30PCL2IN, IDR_SV30PCL2OUT},
				SmartViewTestResource{IDS_STPCL, unittest::parse_all, IDR_SV30PCL3IN, IDR_SV30PCL3OUT},
			}));
		}

		TEST_METHOD(Test_STFBSECURITYDESCRIPTOR)
		{
			test(loadTestData({
				SmartViewTestResource{
					IDS_STFBSECURITYDESCRIPTOR, unittest::parse_all, IDR_SV31FREEBUSYSID1IN, IDR_SV31FREEBUSYSID1OUT},
				SmartViewTestResource{
					IDS_STFBSECURITYDESCRIPTOR, unittest::parse_all, IDR_SV31FREEBUSYSID2IN, IDR_SV31FREEBUSYSID2OUT},
			}));
		}

		TEST_METHOD(Test_STXID)
		{
			test(loadTestData({
				SmartViewTestResource{IDS_STXID, unittest::parse_all, IDR_SV32XID1IN, IDR_SV32XID1OUT},
				SmartViewTestResource{IDS_STXID, unittest::parse_all, IDR_SV32XID2IN, IDR_SV32XID2OUT},
				SmartViewTestResource{IDS_STXID, unittest::parse_all, IDR_SV32XID3IN, IDR_SV32XID3OUT},
				SmartViewTestResource{IDS_STXID, unittest::parse_all, IDR_SV32XID4IN, IDR_SV32XID4OUT},
			}));
		}
	};
} // namespace SmartViewTest