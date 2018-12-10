#include <StdAfx.h>
#include <CppUnitTest.h>
#include <Interpret/Guids.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			template <> inline std::wstring ToString<GUID>(const GUID& t) { return guid::GUIDToString(t); }
		} // namespace CppUnitTestFramework
	} // namespace VisualStudio
} // namespace Microsoft

namespace guidtest
{
	TEST_CLASS(guidtest)
	{
	public:
		// Without this, clang gets weird
		// TODO: figure out why and a better solution
		static const bool dummy_var = true;

		TEST_METHOD(Test_StringToGUID)
		{
			Assert::AreEqual(guid::IID_CAPONE_PROF, guid::StringToGUID(L"{00020D0A-0000-0000-C000-000000000046}"));
			Assert::AreEqual(guid::IID_CAPONE_PROF, guid::StringToGUID(L"00020D0A-0000-0000-C000-000000000046"));
			Assert::AreEqual(guid::IID_CAPONE_PROF, guid::StringToGUID(L"{00020D0A-0000-0000-C000-000000000046"));
			Assert::AreEqual(guid::IID_CAPONE_PROF, guid::StringToGUID(L"00020D0A-0000-0000-C000-000000000046}"));
			Assert::AreEqual(guid::IID_CAPONE_PROF, guid::StringToGUID(L"{0 0 020D0 A-0000-0000-C000-000000000046}  test"));
		}
	};
} // namespace guidtest