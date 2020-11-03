#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/utility/file.h>

namespace filetest
{
	TEST_CLASS(filetest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_GetSystemDirectory)
		{
			// This test is likely not portable
			unittest::AreEqualEx(L"C:\\WINDOWS\\system32", file::GetSystemDirectory());
		}

		TEST_METHOD(Test_ShortenPath)
		{
			unittest::AreEqualEx(
				L"C:\\thisIsAFakePath\\thisIsAnotherFakePath",
				file::ShortenPath(L"C:\\thisIsAFakePath\\thisIsAnotherFakePath"));
			// This test is likely not portable
			unittest::AreEqualEx(L"C:\\PROGRA~1\\COMMON~1\\", file::ShortenPath(L"C:\\Program Files\\Common Files"));
			// This test is likely not portable
			unittest::AreEqualEx(
				L"C:\\PROGRA~2\\COMMON~1\\", file::ShortenPath(L"C:\\Program Files (x86)\\Common Files"));
		}

		TEST_METHOD(Test_BuildFileNameAndPath)
		{
			unittest::AreEqualEx(
				L"c:\\testpath\\subject.xml", file::BuildFileNameAndPath(L".xml", L"subject", L"c:\\testpath\\", NULL));
		}

		TEST_METHOD(Test_GetModuleFileName)
		{
			unittest::AreEqualEx(
				L"C:\\WINDOWS\\SYSTEM32\\ntdll.dll", file::GetModuleFileName(::GetModuleHandleW(L"ntdll.dll")));
		}

		TEST_METHOD(Test_GetFileVersionInfo)
		{
			auto fi1 = file::GetFileVersionInfo(::GetModuleHandleW(NULL));
			unittest::AreEqualEx(L"Microsoft Corporation", fi1[L"CompanyName"]);
			unittest::AreEqualEx(L"© Microsoft Corporation. All rights reserved.", fi1[L"LegalCopyright"]);

			auto fi2 = file::GetFileVersionInfo(::GetModuleHandleW(L"ntdll.dll"));
			unittest::AreEqualEx(L"Microsoft Corporation", fi2[L"CompanyName"]);
			unittest::AreEqualEx(L"NT Layer DLL", fi2[L"FileDescription"]);
			unittest::AreEqualEx(L"Microsoft® Windows® Operating System", fi2[L"ProductName"]);
		}
	};
} // namespace filetest