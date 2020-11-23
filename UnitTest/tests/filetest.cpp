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
			unittest::AreEqualEx(L"c:\\windows\\system32", strings::wstringToLower(file::GetSystemDirectory()));
		}

		TEST_METHOD(Test_ShortenPath)
		{
			unittest::AreEqualEx(L"", file::ShortenPath(L""));
			unittest::AreEqualEx(
				L"c:\\thisisafakepath\\thisisanotherfakepath",
				strings::wstringToLower(file::ShortenPath(L"C:\\thisIsAFakePath\\thisIsAnotherFakePath")));
			// This test is likely not portable
			unittest::AreEqualEx(
				L"c:\\progra~1\\common~1\\",
				strings::wstringToLower(file::ShortenPath(L"C:\\Program Files\\Common Files")));
			// This test is likely not portable
			unittest::AreEqualEx(
				L"c:\\progra~2\\common~1\\",
				strings::wstringToLower(file::ShortenPath(L"C:\\Program Files (x86)\\Common Files")));
		}

		TEST_METHOD(Test_BuildFileNameAndPath)
		{
			unittest::AreEqualEx(
				L"c:\\testpath\\subject.xml", file::BuildFileNameAndPath(L".xml", L"subject", L"c:\\testpath\\", NULL));
			unittest::AreEqualEx(
				L"c:\\testpath\\UnknownSubject.xml", file::BuildFileNameAndPath(L".xml", L"", L"c:\\testpath\\", NULL));
			std::wstring binStr = L"binstr";
			auto bin = SBinary{
				static_cast<ULONG>(binStr.length() * sizeof(WCHAR)),
				reinterpret_cast<LPBYTE>(const_cast<WCHAR*>(binStr.c_str()))};
			unittest::AreEqualEx(
				L"c:\\testpath\\subject_620069006E00730074007200.xml",
				file::BuildFileNameAndPath(L".xml", L"subject", L"c:\\testpath\\", &bin));

			std::string mystringA = std::string(
				"c:\\Now is the time for all good men to come to the aid of the party. This is the way the world "
				"ends. Not with a bang but with a whimper. So long and thanks for all the fish. That's great it "
				"starts with an earthquake, birds and snakes and aeroplanes. Lenny Bruce is not afraid.");
			auto bufA = LPBYTE(mystringA.c_str());
			auto cbA = mystringA.length();
			auto sBinaryA = SBinary{static_cast<ULONG>(cbA), bufA};
			unittest::AreEqualEx(
				L"c:\\testpath\\subject_"
				L"633A5C4E6F77206973207468652074696D6520666F7220616C6C20676F6F64206D656E20746F20636F6D6520746F207468652"
				L"0616964206F66207468652070617274792E2054.xml",
				file::BuildFileNameAndPath(L".xml", L"subject", L"c:\\testpath\\", &sBinaryA));

			unittest::AreEqualEx(
				L"c:\\testpath\\Now is the time for all good men to come to the aid of the party. This is the way the "
				L"world ends. Not with a bang but with a whimper. So long and thanks for all the fish. That's great it "
				L"starts with an earthquake_ birds and snakes an.xml",
				file::BuildFileNameAndPath(
					L".xml",
					L"Now is the time for all good men to come to the aid of the party. This is the way the world "
					L"ends. Not with a bang but with a whimper. So long and thanks for all the fish. That's great it "
					L"starts with an earthquake, birds and snakes and aeroplanes. Lenny Bruce is not afraid.",
					L"c:\\testpath\\",
					NULL));
			unittest::AreEqualEx(
				L"",
				file::BuildFileNameAndPath(
					L".xml",
					L"subject",
					L"c:\\Now is the time for all good men to come to the aid of the party. This is the way the world "
					L"ends. Not with a bang but with a whimper. So long and thanks for all the fish. That's great it "
					L"starts with an earthquake, birds and snakes and aeroplanes. Lenny Bruce is not afraid.",
					NULL));
		}

		TEST_METHOD(Test_GetModuleFileName)
		{
			unittest::AreEqualEx(
				L"c:\\windows\\system32\\ntdll.dll",
				strings::wstringToLower(file::GetModuleFileName(::GetModuleHandleW(L"ntdll.dll"))));
			unittest::AreEqualEx(L"", file::GetModuleFileName(HMODULE(0x1234)));
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