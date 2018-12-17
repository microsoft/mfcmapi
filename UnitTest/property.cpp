#include <StdAfx.h>
#include <CppUnitTest.h>
#include <UnitTest/UnitTest.h>
#include <Property/Property.h>
#include <Property/Attributes.h>
#include <Property/ParseProperty.h>

namespace property
{
	TEST_CLASS(propertyTest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_attribute)
		{
			auto attribute = property::Attribute(L"this", L"that");
			Assert::AreEqual(std::wstring(L"this"), attribute.Key());
			Assert::AreEqual(std::wstring(L"that"), attribute.Value());
			Assert::AreEqual(false, attribute.empty());
			Assert::AreEqual(true, property::Attribute(L"", L"").empty());

			auto copy = property::Attribute(attribute);
			Assert::AreEqual(attribute.toXML(), copy.toXML());
			Assert::AreEqual(attribute.Key(), copy.Key());
			Assert::AreEqual(attribute.Value(), copy.Value());
			Assert::AreEqual(attribute.empty(), copy.empty());
		}

		TEST_METHOD(Test_attributes)
		{
			auto attributes = property::Attributes();
			Assert::AreEqual(std::wstring(L""), attributes.toXML());
			attributes.AddAttribute(L"key", L"value");
			attributes.AddAttribute(L"this", L"that");
			Assert::AreEqual(std::wstring(L" key=\"value\" this=\"that\" "), attributes.toXML());
			Assert::AreEqual(std::wstring(L"that"), attributes.GetAttribute(L"this"));
			Assert::AreEqual(std::wstring(L""), attributes.GetAttribute(L"none"));
		}

		TEST_METHOD(Test_parsing)
		{
			property::Parsing(L"empty", true, property::Attributes{});
			unittest::AreEqualEx(
				std::wstring(L""), property::Parsing(L"", true, property::Attributes{}).toXML(IDS_COLVALUE, 0));

			unittest::AreEqualEx(
				std::wstring(L"test"), property::Parsing(L"test", true, property::Attributes{}).toString());

			auto err = property::Attributes();
			err.AddAttribute(L"err", L"ERROR");
			auto errParsing = property::Parsing(L"myerror", true, err);
			unittest::AreEqualEx(std::wstring(L"Err: ERROR=myerror"), errParsing.toString());

			auto attributes = property::Attributes();
			attributes.AddAttribute(L"key", L"value");
			attributes.AddAttribute(L"this", L"that");
			attributes.AddAttribute(L"cb", L"123");
			auto parsing = property::Parsing(L"parsing", true, attributes);
			unittest::AreEqualEx(
				std::wstring(L"\t<Value key=\"value\" this=\"that\" cb=\"123\" >parsing</Value>\n"),
				parsing.toXML(IDS_COLVALUE, 1));
			unittest::AreEqualEx(std::wstring(L"cb: 123 lpb: parsing"), parsing.toString());

			auto parsing2 = property::Parsing(L"parsing2", false, attributes);
			unittest::AreEqualEx(
				std::wstring(L"<Value key=\"value\" this=\"that\" cb=\"123\" ><![CDATA[parsing2]]></Value>\n"),
				parsing2.toXML(IDS_COLVALUE, 0));
			auto copy = property::Parsing(parsing2);
			unittest::AreEqualEx(parsing2.toXML(IDS_COLVALUE, 0), copy.toXML(IDS_COLVALUE, 0));
			unittest::AreEqualEx(parsing2.toString(), copy.toString());
		}
	};
} // namespace property