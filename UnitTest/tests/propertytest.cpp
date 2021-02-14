#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/property/property.h>
#include <core/property/attributes.h>
#include <core/property/parseProperty.h>

namespace property
{
	// imports for testing
	std::wstring tagopen(_In_ const std::wstring& szTag, int iIndent);
	std::wstring tagclose(_In_ const std::wstring& szTag, int iIndent);
} // namespace property

namespace propertytest
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
			unittest::AreEqualEx(
				std::wstring(L"empty"), property::Parsing(L"empty", true, property::Attributes{}).toString());
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

		TEST_METHOD(Test_property)
		{
			auto prop = property::Property();
			unittest::AreEqualEx(std::wstring(L""), prop.toXML(1));
			unittest::AreEqualEx(std::wstring(L""), prop.toString());
			unittest::AreEqualEx(std::wstring(L""), prop.toAltString());

			const auto parsing = property::Parsing(L"test", true, property::Attributes());
			const auto altparsing = property::Parsing(L"alttest", false, property::Attributes());
			prop.AddParsing(parsing, altparsing);
			unittest::AreEqualEx(
				std::wstring(L"\t<Value>test</Value>\n"
							 L"\t<AltValue><![CDATA[alttest]]></AltValue>\n"),
				prop.toXML(1));
			unittest::AreEqualEx(std::wstring(L"test"), prop.toString());
			unittest::AreEqualEx(std::wstring(L"alttest"), prop.toAltString());

			prop.AddAttribute(L"mv", L"true");
			prop.AddAttribute(L"count", L"2");
			prop.AddMVParsing(prop);
			unittest::AreEqualEx(
				std::wstring(L"\t<Value mv=\"true\" count=\"2\" >\n"
							 L"\t\t<row>\n"
							 L"\t\t\t<Value>test</Value>\n"
							 L"\t\t\t<AltValue><![CDATA[alttest]]></AltValue>\n"
							 L"\t\t</row>\n"
							 L"\t\t<row>\n"
							 L"\t\t\t<Value>test</Value>\n"
							 L"\t\t\t<AltValue><![CDATA[alttest]]></AltValue>\n"
							 L"\t\t</row>\n"
							 L"\t</Value>\n"),
				prop.toXML(1));
			unittest::AreEqualEx(std::wstring(L"2: test; test"), prop.toString());
			unittest::AreEqualEx(std::wstring(L"2: alttest; alttest"), prop.toAltString());
		}

		TEST_METHOD(Test_globals)
		{
			unittest::AreEqualEx(std::wstring(L"\t\t<hello>"), property::tagopen(L"hello", 2));
			unittest::AreEqualEx(std::wstring(L"\t\t\t<hello>"), property::tagopen(L"hello", 3));
			unittest::AreEqualEx(std::wstring(L"\t\t</hello>\n"), property::tagclose(L"hello", 2));
		}

		TEST_METHOD(Test_parseProperty)
		{
			std::wstring szProp;
			std::wstring szAltProp;
			property::parseProperty(nullptr, &szProp, &szAltProp);
			unittest::AreEqualEx(std::wstring(L""), szProp);
			unittest::AreEqualEx(std::wstring(L""), szAltProp);

			auto prop = property::parseProperty(nullptr);
			unittest::AreEqualEx(std::wstring(L""), prop.toXML(0));

			auto sprop = SPropValue{PR_SUBJECT_W, 0, {}};
			sprop.Value.lpszW = const_cast<LPWSTR>(L"hello");
			property::parseProperty(&sprop, &szProp, &szAltProp);
			unittest::AreEqualEx(std::wstring(L"hello"), szProp);
			unittest::AreEqualEx(std::wstring(L"cb: 10 lpb: 680065006C006C006F00"), szAltProp);

			prop = property::parseProperty(&sprop);
			unittest::AreEqualEx(
				std::wstring(L"<Value><![CDATA[hello]]></Value>\n"
							 L"<AltValue cb=\"10\" >680065006C006C006F00</AltValue>\n"),
				prop.toXML(0));
		}
	};
} // namespace propertytest