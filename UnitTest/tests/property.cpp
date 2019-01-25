#include <UnitTest/UnitTest.h>
#include <core/property/property.h>
#include <core/property/attributes.h>

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
			auto attribute = Attribute(L"this", L"that");
			Assert::AreEqual(std::wstring(L"this"), attribute.Key());
			Assert::AreEqual(std::wstring(L"that"), attribute.Value());
			Assert::AreEqual(false, attribute.empty());
			Assert::AreEqual(true, Attribute(L"", L"").empty());

			auto copy = Attribute(attribute);
			Assert::AreEqual(attribute.toXML(), copy.toXML());
			Assert::AreEqual(attribute.Key(), copy.Key());
			Assert::AreEqual(attribute.Value(), copy.Value());
			Assert::AreEqual(attribute.empty(), copy.empty());
		}

		TEST_METHOD(Test_attributes)
		{
			auto attributes = Attributes();
			Assert::AreEqual(std::wstring(L""), attributes.toXML());
			attributes.AddAttribute(L"key", L"value");
			attributes.AddAttribute(L"this", L"that");
			Assert::AreEqual(std::wstring(L" key=\"value\" this=\"that\" "), attributes.toXML());
			Assert::AreEqual(std::wstring(L"that"), attributes.GetAttribute(L"this"));
			Assert::AreEqual(std::wstring(L""), attributes.GetAttribute(L"none"));
		}

		TEST_METHOD(Test_parsing)
		{
			Parsing(L"empty", true, Attributes{});
			unittest::AreEqualEx(std::wstring(L""), Parsing(L"", true, Attributes{}).toXML(IDS_COLVALUE, 0));

			unittest::AreEqualEx(std::wstring(L"test"), Parsing(L"test", true, Attributes{}).toString());

			auto err = Attributes();
			err.AddAttribute(L"err", L"ERROR");
			auto errParsing = Parsing(L"myerror", true, err);
			unittest::AreEqualEx(std::wstring(L"Err: ERROR=myerror"), errParsing.toString());

			auto attributes = Attributes();
			attributes.AddAttribute(L"key", L"value");
			attributes.AddAttribute(L"this", L"that");
			attributes.AddAttribute(L"cb", L"123");
			auto parsing = Parsing(L"parsing", true, attributes);
			unittest::AreEqualEx(
				std::wstring(L"\t<Value key=\"value\" this=\"that\" cb=\"123\" >parsing</Value>\n"),
				parsing.toXML(IDS_COLVALUE, 1));
			unittest::AreEqualEx(std::wstring(L"cb: 123 lpb: parsing"), parsing.toString());

			auto parsing2 = Parsing(L"parsing2", false, attributes);
			unittest::AreEqualEx(
				std::wstring(L"<Value key=\"value\" this=\"that\" cb=\"123\" ><![CDATA[parsing2]]></Value>\n"),
				parsing2.toXML(IDS_COLVALUE, 0));
			auto copy = Parsing(parsing2);
			unittest::AreEqualEx(parsing2.toXML(IDS_COLVALUE, 0), copy.toXML(IDS_COLVALUE, 0));
			unittest::AreEqualEx(parsing2.toString(), copy.toString());
		}

		TEST_METHOD(Test_property)
		{
			auto prop = Property();
			unittest::AreEqualEx(std::wstring(L""), prop.toXML(1));
			unittest::AreEqualEx(std::wstring(L""), prop.toString());
			unittest::AreEqualEx(std::wstring(L""), prop.toAltString());

			const auto parsing = Parsing(L"test", true, Attributes());
			const auto altparsing = Parsing(L"alttest", false, Attributes());
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
							 L"\t\t\t<AltValue>test</AltValue>\n"
							 L"\t\t</row>\n"
							 L"\t\t<row>\n"
							 L"\t\t\t<Value>test</Value>\n"
							 L"\t\t\t<AltValue>test</AltValue>\n"
							 L"\t\t</row>\n"
							 L"\t</Value>\n"),
				prop.toXML(1));
			unittest::AreEqualEx(std::wstring(L"2: test; test"), prop.toString());
			unittest::AreEqualEx(std::wstring(L"2: alttest; alttest"), prop.toAltString());
		}
	};
} // namespace property