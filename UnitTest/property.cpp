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
			Assert::AreEqual(std::wstring(L"this=\"that\" "), copy.toXML());
		}
	};
} // namespace property