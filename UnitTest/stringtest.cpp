#include "stdafx.h"
#include "CppUnitTest.h"
#include "Interpret\String.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

// In case we have any tests that depend on MFC being loaded, here's how to load it
// AfxWinInit(::GetModuleHandleW(L"UnitTest.dll"), nullptr, ::GetCommandLine(), 0);

namespace stringtest
{
	TEST_CLASS(stringtest)
	{
	public:
		TEST_METHOD(Test_formatmessagesys)
		{
			Assert::AreEqual(
				wstring(L"No service is operating at the destination network endpoint on the remote system.\r\n"),
				formatmessagesys(1234));

			Assert::AreEqual(
				wstring(L"The operation completed successfully.\r\n"),
				formatmessagesys(0));

			Assert::AreEqual(
				wstring(L""),
				formatmessagesys(0xFFFFFFFF));
		}

		TEST_METHOD(Test_loadstring)
		{
			// loadstring has two flavors:
			// wstring loadstring(HINSTANCE hInstance, DWORD dwID);
			// wstring loadstring(DWORD dwID);
			// The latter calls the former with a null hInstance.
			// This will try to load the string from the executable, which is fine for MFCMAPI and MrMAPI
			// In our unit tests, we must load strings from UnitTest.dll, so we can only test
			// the variant which accepts an HINSTANCE
			auto hInstance = ::GetModuleHandleW(L"UnitTest.dll");

			// A resource which does not exist
			Assert::AreEqual(
				wstring(L""),
				loadstring(hInstance, 1234));

			// A resource which does exist
			Assert::AreEqual(
				wstring(L"\r\n\tUnknown Data = "),
				loadstring(hInstance, IDS_EXTENDEDFLAGUNKNOWN));

			// A resource which does exist, but loaded from the wrong dll
			Assert::AreNotEqual(
				wstring(L"\r\n\tUnknown Data = "),
				loadstring(nullptr, IDS_EXTENDEDFLAGUNKNOWN));
		}

		TEST_METHOD(Test_format)
		{
			// Since format is a passthrough to formatV, this will also cover formatV
			Assert::AreEqual(wstring(L"Hello"), format(L"Hello"));
			Assert::AreEqual(wstring(L"Hello world"), format(L"Hello %hs", "world"));
			Assert::AreEqual(wstring(L"Hello (null)"), format(L"Hello %hs"));
		}

		TEST_METHOD(Test_formatmessage)
		{
			// Also tests formatmessageV
			Assert::AreEqual(
				wstring(L"Hello world"),
				formatmessage(L"%1!hs! %2", "Hello", L"world"));
		}

		TEST_METHOD(Test_stringConverters)
		{
			LPCTSTR lpctstr = _T("Hello World");
			LPCSTR lpcstr = "Hello World";
			auto tstr = tstring(lpctstr);
			auto wstr = wstring(L"Hello World");
			auto str = string(lpcstr);
			auto wstrLower = wstring(L"hello world");

			Assert::AreEqual(tstr, wstringTotstring(wstr));
			Assert::AreEqual(str, wstringTostring(wstr));
			Assert::AreEqual(wstr, stringTowstring(str));
			Assert::AreEqual(wstr, LPCTSTRToWstring(lpctstr));
			Assert::AreEqual(wstr, LPCSTRToWstring(lpcstr));
			Assert::AreEqual(wstrLower, wstringToLower(wstr));

			Assert::AreEqual(wstring(L""), LPCTSTRToWstring(nullptr));
			Assert::AreEqual(wstring(L""), LPCSTRToWstring(nullptr));
		}

		TEST_METHOD(Test_wstringToUlong)
		{
			Assert::AreEqual(ULONG(0), wstringToUlong(L"", 10));
			Assert::AreEqual(ULONG(1234), wstringToUlong(L"1234", 10));
			Assert::AreEqual(ULONG(1234), wstringToUlong(wstring(L"1234"), 10));
			Assert::AreEqual(ULONG(1234), wstringToUlong(L"1234test", 10, false));
			Assert::AreEqual(ULONG(0), wstringToUlong(L"1234test", 10));
			Assert::AreEqual(ULONG(0), wstringToUlong(L"1234test", 10, true));
			Assert::AreEqual(ULONG(4660), wstringToUlong(L"1234", 16));
			Assert::AreEqual(ULONG(4660), wstringToUlong(L"0x1234", 16));
			Assert::AreEqual(ULONG(0), wstringToUlong(L"x1234", 16));
			Assert::AreEqual(ULONG(42), wstringToUlong(L"101010", 2));
			Assert::AreEqual(ULONG(42), wstringToUlong(L"1010102", 2, false));

			Assert::AreEqual(ULONG_MAX, wstringToUlong(L"-1", 10));
			Assert::AreEqual(ULONG_MAX - 1, wstringToUlong(L"-2", 10));
			Assert::AreEqual(ULONG_MAX - 1, wstringToUlong(L"4294967294", 10));
			Assert::AreEqual(ULONG_MAX, wstringToUlong(L"4294967295", 10));
			Assert::AreEqual(ULONG_MAX, wstringToUlong(L"4294967296", 10));
		}

		TEST_METHOD(Test_wstringToLong)
		{
			Assert::AreEqual(long(0), wstringToLong(L"", 10));
			Assert::AreEqual(long(1234), wstringToLong(L"1234", 10));
			Assert::AreEqual(long(1234), wstringToLong(wstring(L"1234"), 10));
			Assert::AreEqual(long(0), wstringToLong(L"1234test", 10));
			Assert::AreEqual(long(4660), wstringToLong(L"1234", 16));
			Assert::AreEqual(long(4660), wstringToLong(L"0x1234", 16));
			Assert::AreEqual(long(0), wstringToLong(L"x1234", 16));
			Assert::AreEqual(long(42), wstringToLong(L"101010", 2));
			Assert::AreEqual(long(0), wstringToLong(L"1010102", 2));

			Assert::AreEqual(long(-1), wstringToLong(L"-1", 10));
			Assert::AreEqual(long(-2), wstringToLong(L"-2", 10));
			Assert::AreEqual(LONG_MAX - 1, wstringToLong(L"2147483646", 10));
			Assert::AreEqual(LONG_MAX, wstringToLong(L"2147483647", 10));
			Assert::AreEqual(LONG_MAX, wstringToLong(L"2147483648", 10));
		}

		TEST_METHOD(Test_wstringToDouble)
		{
			Assert::AreEqual(double(1234), wstringToDouble(L"1234"));
			Assert::AreEqual(double(12.34), wstringToDouble(L"12.34"));
			Assert::AreEqual(double(0), wstringToDouble(L"12.34test"));
			Assert::AreEqual(double(1e+52), wstringToDouble(L"9999999999999999999999999999999999999999999999999999"));
			Assert::AreEqual(double(1e-52), wstringToDouble(L".0000000000000000000000000000000000000000000000000001"));
			Assert::AreEqual(double(0), wstringToDouble(L"0"));
			Assert::AreEqual(double(0), wstringToDouble(L""));
		}

		TEST_METHOD(Test_wstringToInt64)
		{
			Assert::AreEqual(__int64(1234), wstringToInt64(L"1234"));
			Assert::AreEqual(__int64(-1), wstringToInt64(L"-1"));
			Assert::AreEqual(INT64_MAX - 1, wstringToInt64(L"9223372036854775806"));
			Assert::AreEqual(INT64_MAX, wstringToInt64(L"9223372036854775807"));
			Assert::AreEqual(INT64_MAX, wstringToInt64(L"9223372036854775808"));
			Assert::AreEqual(__int64(0), wstringToInt64(L"0"));
			Assert::AreEqual(__int64(0), wstringToInt64(L""));
		}

		TEST_METHOD(Test_StripCharacter)
		{
			Assert::AreEqual(wstring(L"1245"), StripCharacter(L"12345", L'3'));
			Assert::AreEqual(wstring(L"12345"), StripCharacter(L"12345", L'6'));
			Assert::AreEqual(wstring(L"12345"), StripCharacter(L" 1 2 3 4 5", L' '));
			const wstring conststring = wstring(L"12345");
			Assert::AreEqual(wstring(L"1345"), StripCharacter(conststring, L'2'));
		}

		TEST_METHOD(Test_StripCarriage)
		{
			Assert::AreEqual(wstring(L"12345"), StripCarriage(L"1\r2345\r\r"));
		}

		TEST_METHOD(Test_CleanString)
		{
			Assert::AreEqual(wstring(L"12345"), CleanString(L"1\r2345\r\r"));
			Assert::AreEqual(wstring(L"12345"), CleanString(L"1\r23\n\r\n45\r\n\r"));
		}

		TEST_METHOD(Test_TrimString)
		{
			Assert::AreEqual(wstring(L"12345"), TrimString(L"12345"));
			Assert::AreEqual(wstring(L"12345"), TrimString(L"    12345"));
			Assert::AreEqual(wstring(L""), TrimString(L"  "));
		}

		TEST_METHOD(Test_ScrubStringForXML)
		{
			Assert::AreEqual(wstring(L"12345"), ScrubStringForXML(L"12345"));
			Assert::AreEqual(wstring(L"1."), ScrubStringForXML(L"1\1"));
			Assert::AreEqual(wstring(L"2."), ScrubStringForXML(L"2\2"));
			Assert::AreEqual(wstring(L"3."), ScrubStringForXML(L"3\3"));
			Assert::AreEqual(wstring(L"4."), ScrubStringForXML(L"4\4"));
			Assert::AreEqual(wstring(L"5."), ScrubStringForXML(L"5\5"));
			Assert::AreEqual(wstring(L"6."), ScrubStringForXML(L"6\6"));
			Assert::AreEqual(wstring(L"7."), ScrubStringForXML(L"7\7"));
			Assert::AreEqual(wstring(L"7a."), ScrubStringForXML(L"7a\a"));
			Assert::AreEqual(wstring(L"8."), ScrubStringForXML(L"8\010"));
			Assert::AreEqual(wstring(L"8a."), ScrubStringForXML(L"8a\x08"));
			Assert::AreEqual(wstring(L"8b."), ScrubStringForXML(L"8b\f"));
			Assert::AreEqual(wstring(L"9\t"), ScrubStringForXML(L"9\x09"));
			Assert::AreEqual(wstring(L"A\n"), ScrubStringForXML(L"A\x0A"));
			Assert::AreEqual(wstring(L"B."), ScrubStringForXML(L"B\x0B"));
			Assert::AreEqual(wstring(L"Ba."), ScrubStringForXML(L"Ba\b"));
			Assert::AreEqual(wstring(L"C."), ScrubStringForXML(L"C\x0C"));
			Assert::AreEqual(wstring(L"D\r"), ScrubStringForXML(L"D\x0D"));
			Assert::AreEqual(wstring(L"E."), ScrubStringForXML(L"E\x0E"));
			Assert::AreEqual(wstring(L"F."), ScrubStringForXML(L"F\x0F"));
			Assert::AreEqual(wstring(L"10."), ScrubStringForXML(L"10\x10"));
			Assert::AreEqual(wstring(L"11."), ScrubStringForXML(L"11\x11"));
			Assert::AreEqual(wstring(L"12."), ScrubStringForXML(L"12\x12"));
			Assert::AreEqual(wstring(L"13."), ScrubStringForXML(L"13\x13"));
			Assert::AreEqual(wstring(L"14."), ScrubStringForXML(L"14\x14"));
			Assert::AreEqual(wstring(L"15."), ScrubStringForXML(L"15\x15"));
			Assert::AreEqual(wstring(L"16."), ScrubStringForXML(L"16\x16"));
			Assert::AreEqual(wstring(L"17."), ScrubStringForXML(L"17\x17"));
			Assert::AreEqual(wstring(L"18."), ScrubStringForXML(L"18\x18"));
			Assert::AreEqual(wstring(L"19."), ScrubStringForXML(L"19\x19"));
			Assert::AreEqual(wstring(L"20 "), ScrubStringForXML(L"20\x20"));
		}

		TEST_METHOD(Test_SanitizeFileName)
		{
			Assert::AreEqual(wstring(L"_This_ is_a_bad string_!_"), SanitizeFileName(L"[This] is\ra\nbad string?!?"));
			Assert::AreEqual(wstring(L"____________________"), SanitizeFileName(L"^&*-+=[]\\|;:\",<>/?\r\n"));
		}

		TEST_METHOD(Test_indent)
		{
			Assert::AreEqual(wstring(L""), indent(0));
			Assert::AreEqual(wstring(L"\t"), indent(1));
			Assert::AreEqual(wstring(L"\t\t\t\t\t"), indent(5));
		}

		wstring mystringW = wstring(L"mystring");
		LPBYTE bufW = (LPBYTE)mystringW.c_str();
		size_t cbW = mystringW.length() * sizeof(WCHAR);
		SBinary sBinaryW = SBinary{ (ULONG)cbW, bufW };
		vector<BYTE> myStringWvector = vector<BYTE>(bufW, bufW + cbW);

		string mystringA = string("mystring");
		LPBYTE bufA = (LPBYTE)mystringA.c_str();
		size_t cbA = mystringA.length();
		SBinary sBinaryA = SBinary{ (ULONG)cbA, bufA };
		vector<BYTE> myStringAvector = vector<BYTE>(bufA, bufA + cbA);

		vector<BYTE> vector_abcdW = vector<BYTE>({ 0x61, 0, 0x62, 0, 0x63, 0, 0x64, 0 });
		vector<BYTE> vector_abNULLdW = vector<BYTE>({ 0x61, 0, 0x62, 0, 0x00, 0, 0x64, 0 });
		vector<BYTE> vector_tabcrlfW = vector<BYTE>({ 0x9, 0, 0xa, 0, 0xd, 0 });

		vector<BYTE> vector_abcdA = vector<BYTE>({ 0x61, 0x62, 0x63, 0x64 });
		vector<BYTE> vector_abNULLdA = vector<BYTE>({ 0x61, 0x62, 0x00, 0x64 });
		vector<BYTE> vector_tabcrlfA = vector<BYTE>({ 0x9, 0xa, 0xd });

		TEST_METHOD(Test_BinToTextStringW)
		{
			Assert::AreEqual(wstring(L""), BinToTextStringW(0, true));
			Assert::AreEqual(wstring(L""), BinToTextStringW(0, false));
			Assert::AreEqual(wstring(L""), BinToTextStringW(nullptr, true));

			Assert::AreEqual(mystringW, BinToTextStringW(&sBinaryW, false));
			Assert::AreEqual(mystringW, BinToTextStringW(myStringWvector, false));
			Assert::AreEqual(wstring(L"abcd"), BinToTextStringW(vector_abcdW, false));
			Assert::AreEqual(wstring(L"ab.d"), BinToTextStringW(vector_abNULLdW, false));
			Assert::AreEqual(wstring(L"\t\n\r"), BinToTextStringW(vector_tabcrlfW, true));
			Assert::AreEqual(wstring(L"..."), BinToTextStringW(vector_tabcrlfW, false));
		}

		TEST_METHOD(Test_BinToTextString)
		{
			Assert::AreEqual(wstring(L""), BinToTextString(0, true));
			Assert::AreEqual(wstring(L""), BinToTextString(0, false));
			Assert::AreEqual(wstring(L""), BinToTextString(nullptr, true));

			Assert::AreEqual(mystringW, BinToTextString(&sBinaryA, false));
			auto sBinary = SBinary{ (ULONG)vector_abcdA.size(), vector_abcdA.data() };
			Assert::AreEqual(wstring(L"abcd"), BinToTextString(&sBinary, false));
			sBinary = SBinary{ (ULONG)vector_abNULLdA.size(), vector_abNULLdA.data() };
			Assert::AreEqual(wstring(L"ab.d"), BinToTextString(&sBinary, false));
			sBinary = SBinary{ (ULONG)vector_tabcrlfA.size(), vector_tabcrlfA.data() };
			Assert::AreEqual(wstring(L"\t\n\r"), BinToTextString(&sBinary, true));
			sBinary = SBinary{ (ULONG)vector_tabcrlfA.size(), vector_tabcrlfA.data() };
			Assert::AreEqual(wstring(L"..."), BinToTextString(&sBinary, false));
		}

		TEST_METHOD(Test_BinToHexString)
		{
			Assert::AreEqual(wstring(L"NULL"), BinToHexString(vector<BYTE>(), false));
			Assert::AreEqual(wstring(L"FF"), BinToHexString(vector<BYTE>{0xFF}, false));
			Assert::AreEqual(wstring(L""), BinToHexString(nullptr, false));
			Assert::AreEqual(wstring(L"6D79737472696E67"), BinToHexString(&sBinaryA, false));
			Assert::AreEqual(wstring(L"6D00790073007400720069006E006700"), BinToHexString(&sBinaryW, false));
			Assert::AreEqual(wstring(L"cb: 8 lpb: 6D79737472696E67"), BinToHexString(vector<BYTE>(bufA, bufA + cbA), true));
			Assert::AreEqual(wstring(L"6100620063006400"), BinToHexString(vector_abcdW, false));
			Assert::AreEqual(wstring(L"6100620000006400"), BinToHexString(vector_abNULLdW, false));
			Assert::AreEqual(wstring(L"cb: 6 lpb: 09000A000D00"), BinToHexString(vector_tabcrlfW, true));
			Assert::AreEqual(wstring(L"09000A000D00"), BinToHexString(vector_tabcrlfW, false));
		}

		TEST_METHOD(Test_HexStringToBin)
		{
			Assert::AreEqual(vector<BYTE>(), HexStringToBin(wstring(L"12345")));
			Assert::AreEqual(vector_abcdW, HexStringToBin(wstring(L"6100620063006400")));
			Assert::AreEqual(vector_abcdW, HexStringToBin(wstring(L"0x6100620063006400")));
			Assert::AreEqual(vector_abcdW, HexStringToBin(wstring(L"0X6100620063006400")));
			Assert::AreEqual(vector_abcdW, HexStringToBin(wstring(L"x6100620063006400")));
			Assert::AreEqual(vector_abcdW, HexStringToBin(wstring(L"X6100620063006400")));
			Assert::AreEqual(vector_abNULLdW, HexStringToBin(wstring(L"6100620000006400")));
			Assert::AreEqual(vector_tabcrlfW, HexStringToBin(wstring(L"09000A000D00")));
			Assert::AreEqual(myStringWvector, HexStringToBin(wstring(L"6D00790073007400720069006E006700")));
		}

		void ByteVectorToLPBYTETest(const vector<BYTE>& bin)
		{
			auto bytes = ByteVectorToLPBYTE(bin);
			for (size_t i = 0; i < bin.size(); i++)
			{
				Assert::AreEqual(bytes[i], bin[i]);
			}

			delete[] bytes;
		}

		TEST_METHOD(Test_ByteVectorToLPBYTE)
		{
			ByteVectorToLPBYTETest(vector<BYTE>{});
			ByteVectorToLPBYTETest(vector_abcdW);
			ByteVectorToLPBYTETest(vector_abNULLdW);
			ByteVectorToLPBYTETest(vector_tabcrlfW);

			ByteVectorToLPBYTETest(vector_abcdA);
			ByteVectorToLPBYTETest(vector_abNULLdA);
			ByteVectorToLPBYTETest(vector_tabcrlfA);
		}

		wstring fullstring = wstring(L"this is a string  yes");
		vector<wstring> words = { wstring(L"this"), wstring(L"is"), wstring(L"a"), wstring(L"string"), wstring(L""), wstring(L"yes") };

		TEST_METHOD(Test_split)
		{
			auto splitWords = split(fullstring, L' ');

			size_t i = 0;
			for (i = 0; i < splitWords.size(); i++)
			{
				Assert::IsTrue(i < words.size());
				Assert::AreEqual(words[i], splitWords[i]);
			}

			Assert::IsTrue(i == words.size());
		}

		TEST_METHOD(Test_join)
		{
			Assert::AreEqual(fullstring, join(words, L' '));
		}
	};
}

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			template<> inline std::wstring ToString<__int64>(const __int64& t) { RETURN_WIDE_STRING(t); }
			template<> inline std::wstring ToString<vector<BYTE>>(const vector<BYTE>& t) { RETURN_WIDE_STRING(t.data()); }
		}
	}
}