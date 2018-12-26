#include <../StdAfx.h>
#include <CppUnitTest.h>
#include <UnitTest/UnitTest.h>
#include <Interpret/String.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace stringtest
{
	TEST_CLASS(stringtest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_formatmessagesys)
		{
			Assert::AreEqual(
				std::wstring(L"No service is operating at the destination network endpoint on the remote system.\r\n"),
				strings::formatmessagesys(1234));

			Assert::AreEqual(std::wstring(L"The operation completed successfully.\r\n"), strings::formatmessagesys(0));

			Assert::AreEqual(std::wstring(L""), strings::formatmessagesys(0xFFFFFFFF));
		}

		TEST_METHOD(Test_loadstring)
		{
			// A resource which does not exist
			Assert::AreEqual(std::wstring(L""), strings::loadstring(1234));

			// A resource which does exist
			Assert::AreEqual(std::wstring(L"Flags: "), strings::loadstring(IDS_FLAGS_PREFIX));
		}

		TEST_METHOD(Test_format)
		{
			// Since format is a passthrough to formatV, this will also cover formatV
			Assert::AreEqual(std::wstring(L"Hello"), strings::format(L"Hello"));
			Assert::AreEqual(std::wstring(L"Hello world"), strings::format(L"Hello %hs", "world"));
			Assert::AreEqual(std::wstring(L""), strings::format(L"", 1, 2));
		}

		TEST_METHOD(Test_formatmessage)
		{
			// Also tests formatmessageV
			Assert::AreEqual(std::wstring(L"Hello world"), strings::formatmessage(L"%1!hs! %2", "Hello", L"world"));
			Assert::AreEqual(std::wstring(L""), strings::formatmessage(L"", 1, 2));
			Assert::AreEqual(std::wstring(L"test"), strings::formatmessage(IDS_TITLEBARPLAIN, L"test"));
		}

		TEST_METHOD(Test_stringConverters)
		{
			const auto lpctstr = _T("Hello World");
			const auto lpcstr = "Hello World";
			const auto tstr = strings::tstring(lpctstr);
			const auto wstr = std::wstring(L"Hello World");
			const auto str = std::string(lpcstr);
			const auto wstrLower = std::wstring(L"hello world");

			Assert::AreEqual(tstr, strings::wstringTotstring(wstr));
			Assert::AreEqual(str, strings::wstringTostring(wstr));
			Assert::AreEqual(wstr, strings::stringTowstring(str));
			Assert::AreEqual(wstr, strings::LPCTSTRToWstring(lpctstr));
			Assert::AreEqual(wstr, strings::LPCSTRToWstring(lpcstr));
			Assert::AreEqual(wstrLower, strings::wstringToLower(wstr));
			Assert::AreEqual(
				std::wstring(L"abc\xDC\xA7\x40\xC8\xC0\x42"), strings::stringTowstring("abc\xDC\xA7\x40\xC8\xC0\x42"));

			Assert::AreEqual(std::wstring(L""), strings::LPCTSTRToWstring(nullptr));
			Assert::AreEqual(std::wstring(L""), strings::LPCSTRToWstring(nullptr));
			Assert::AreEqual(L"test", strings::wstringToLPCWSTR(L"test"));
			Assert::AreEqual(L"", strings::wstringToLPCWSTR(L""));
		}

		TEST_METHOD(Test_wstringToUlong)
		{
			Assert::AreEqual(ULONG(0), strings::wstringToUlong(L"", 10));
			Assert::AreEqual(ULONG(1234), strings::wstringToUlong(L"1234", 10));
			Assert::AreEqual(ULONG(1234), strings::wstringToUlong(std::wstring(L"1234"), 10));
			Assert::AreEqual(ULONG(1234), strings::wstringToUlong(L"1234test", 10, false));
			Assert::AreEqual(ULONG(0), strings::wstringToUlong(L"1234test", 10));
			Assert::AreEqual(ULONG(0), strings::wstringToUlong(L"1234test", 10, true));
			Assert::AreEqual(ULONG(4660), strings::wstringToUlong(L"1234", 16));
			Assert::AreEqual(ULONG(4660), strings::wstringToUlong(L"0x1234", 16));
			Assert::AreEqual(ULONG(0), strings::wstringToUlong(L"x1234", 16));
			Assert::AreEqual(ULONG(42), strings::wstringToUlong(L"101010", 2));
			Assert::AreEqual(ULONG(42), strings::wstringToUlong(L"1010102", 2, false));

			Assert::AreEqual(ULONG_MAX, strings::wstringToUlong(L"-1", 10));
			Assert::AreEqual(ULONG_MAX - 1, strings::wstringToUlong(L"-2", 10));
			Assert::AreEqual(ULONG_MAX - 1, strings::wstringToUlong(L"4294967294", 10));
			Assert::AreEqual(ULONG_MAX, strings::wstringToUlong(L"4294967295", 10));
			Assert::AreEqual(ULONG_MAX, strings::wstringToUlong(L"4294967296", 10));
		}

		TEST_METHOD(Test_wstringToLong)
		{
			Assert::AreEqual(long(0), strings::wstringToLong(L"", 10));
			Assert::AreEqual(long(1234), strings::wstringToLong(L"1234", 10));
			Assert::AreEqual(long(1234), strings::wstringToLong(std::wstring(L"1234"), 10));
			Assert::AreEqual(long(0), strings::wstringToLong(L"1234test", 10));
			Assert::AreEqual(long(4660), strings::wstringToLong(L"1234", 16));
			Assert::AreEqual(long(4660), strings::wstringToLong(L"0x1234", 16));
			Assert::AreEqual(long(0), strings::wstringToLong(L"x1234", 16));
			Assert::AreEqual(long(42), strings::wstringToLong(L"101010", 2));
			Assert::AreEqual(long(0), strings::wstringToLong(L"1010102", 2));

			Assert::AreEqual(long(-1), strings::wstringToLong(L"-1", 10));
			Assert::AreEqual(long(-2), strings::wstringToLong(L"-2", 10));
			Assert::AreEqual(LONG_MAX - 1, strings::wstringToLong(L"2147483646", 10));
			Assert::AreEqual(LONG_MAX, strings::wstringToLong(L"2147483647", 10));
			Assert::AreEqual(LONG_MAX, strings::wstringToLong(L"2147483648", 10));
		}

		TEST_METHOD(Test_wstringToDouble)
		{
			Assert::AreEqual(double(1234), strings::wstringToDouble(L"1234"));
			Assert::AreEqual(double(12.34), strings::wstringToDouble(L"12.34"));
			Assert::AreEqual(double(0), strings::wstringToDouble(L"12.34test"));
			Assert::AreEqual(
				double(1e+52), strings::wstringToDouble(L"9999999999999999999999999999999999999999999999999999"));
			Assert::AreEqual(
				double(1e-52), strings::wstringToDouble(L".0000000000000000000000000000000000000000000000000001"));
			Assert::AreEqual(double(0), strings::wstringToDouble(L"0"));
			Assert::AreEqual(double(0), strings::wstringToDouble(L""));
		}

		TEST_METHOD(Test_wstringToInt64)
		{
			Assert::AreEqual(__int64(1234), strings::wstringToInt64(L"1234"));
			Assert::AreEqual(__int64(-1), strings::wstringToInt64(L"-1"));
			Assert::AreEqual(INT64_MAX - 1, strings::wstringToInt64(L"9223372036854775806"));
			Assert::AreEqual(INT64_MAX, strings::wstringToInt64(L"9223372036854775807"));
			Assert::AreEqual(INT64_MAX, strings::wstringToInt64(L"9223372036854775808"));
			Assert::AreEqual(__int64(0), strings::wstringToInt64(L"0"));
			Assert::AreEqual(__int64(0), strings::wstringToInt64(L""));
		}

		TEST_METHOD(Test_StripCharacter)
		{
			Assert::AreEqual(std::wstring(L"1245"), strings::StripCharacter(L"12345", L'3'));
			Assert::AreEqual(std::wstring(L"12345"), strings::StripCharacter(L"12345", L'6'));
			Assert::AreEqual(std::wstring(L"12345"), strings::StripCharacter(L" 1 2 3 4 5", L' '));
			const std::wstring conststring = L"12345";
			Assert::AreEqual(std::wstring(L"1345"), strings::StripCharacter(conststring, L'2'));
		}

		TEST_METHOD(Test_StripCarriage)
		{
			Assert::AreEqual(std::wstring(L"12345"), strings::StripCarriage(L"1\r2345\r\r"));
		}

		TEST_METHOD(Test_StripCRLF)
		{
			Assert::AreEqual(std::wstring(L"12345"), strings::StripCRLF(L"1\r2345\r\r"));
			Assert::AreEqual(std::wstring(L"12345"), strings::StripCRLF(L"1\r23\n\r\n45\r\n\r"));
		}

		TEST_METHOD(Test_TrimString)
		{
			Assert::AreEqual(std::wstring(L"12345"), strings::trim(L"12345"));
			Assert::AreEqual(std::wstring(L"12345"), strings::trim(L"    12345"));
			Assert::AreEqual(std::wstring(L""), strings::trim(L"  "));
		}

		TEST_METHOD(Test_ScrubStringForXML)
		{
			Assert::AreEqual(std::wstring(L"12345"), strings::ScrubStringForXML(L"12345"));
			Assert::AreEqual(std::wstring(L"1."), strings::ScrubStringForXML(L"1\1"));
			Assert::AreEqual(std::wstring(L"2."), strings::ScrubStringForXML(L"2\2"));
			Assert::AreEqual(std::wstring(L"3."), strings::ScrubStringForXML(L"3\3"));
			Assert::AreEqual(std::wstring(L"4."), strings::ScrubStringForXML(L"4\4"));
			Assert::AreEqual(std::wstring(L"5."), strings::ScrubStringForXML(L"5\5"));
			Assert::AreEqual(std::wstring(L"6."), strings::ScrubStringForXML(L"6\6"));
			Assert::AreEqual(std::wstring(L"7."), strings::ScrubStringForXML(L"7\7"));
			Assert::AreEqual(std::wstring(L"7a."), strings::ScrubStringForXML(L"7a\a"));
			Assert::AreEqual(std::wstring(L"8."), strings::ScrubStringForXML(L"8\010"));
			Assert::AreEqual(std::wstring(L"8a."), strings::ScrubStringForXML(L"8a\x08"));
			Assert::AreEqual(std::wstring(L"8b."), strings::ScrubStringForXML(L"8b\f"));
			Assert::AreEqual(std::wstring(L"9\t"), strings::ScrubStringForXML(L"9\x09"));
			Assert::AreEqual(std::wstring(L"A\n"), strings::ScrubStringForXML(L"A\x0A"));
			Assert::AreEqual(std::wstring(L"B."), strings::ScrubStringForXML(L"B\x0B"));
			Assert::AreEqual(std::wstring(L"Ba."), strings::ScrubStringForXML(L"Ba\b"));
			Assert::AreEqual(std::wstring(L"C."), strings::ScrubStringForXML(L"C\x0C"));
			Assert::AreEqual(std::wstring(L"D\r"), strings::ScrubStringForXML(L"D\x0D"));
			Assert::AreEqual(std::wstring(L"E."), strings::ScrubStringForXML(L"E\x0E"));
			Assert::AreEqual(std::wstring(L"F."), strings::ScrubStringForXML(L"F\x0F"));
			Assert::AreEqual(std::wstring(L"10."), strings::ScrubStringForXML(L"10\x10"));
			Assert::AreEqual(std::wstring(L"11."), strings::ScrubStringForXML(L"11\x11"));
			Assert::AreEqual(std::wstring(L"12."), strings::ScrubStringForXML(L"12\x12"));
			Assert::AreEqual(std::wstring(L"13."), strings::ScrubStringForXML(L"13\x13"));
			Assert::AreEqual(std::wstring(L"14."), strings::ScrubStringForXML(L"14\x14"));
			Assert::AreEqual(std::wstring(L"15."), strings::ScrubStringForXML(L"15\x15"));
			Assert::AreEqual(std::wstring(L"16."), strings::ScrubStringForXML(L"16\x16"));
			Assert::AreEqual(std::wstring(L"17."), strings::ScrubStringForXML(L"17\x17"));
			Assert::AreEqual(std::wstring(L"18."), strings::ScrubStringForXML(L"18\x18"));
			Assert::AreEqual(std::wstring(L"19."), strings::ScrubStringForXML(L"19\x19"));
			Assert::AreEqual(std::wstring(L"20 "), strings::ScrubStringForXML(L"20\x20"));
		}

		TEST_METHOD(Test_SanitizeFileName)
		{
			Assert::AreEqual(
				std::wstring(L"_This_ is_a_bad string_!_"), strings::SanitizeFileName(L"[This] is\ra\nbad string?!?"));
			Assert::AreEqual(
				std::wstring(L"____________________"), strings::SanitizeFileName(L"^&*-+=[]\\|;:\",<>/?\r\n"));
		}

		TEST_METHOD(Test_indent)
		{
			Assert::AreEqual(std::wstring(L""), strings::indent(0));
			Assert::AreEqual(std::wstring(L"\t"), strings::indent(1));
			Assert::AreEqual(std::wstring(L"\t\t\t\t\t"), strings::indent(5));
		}

		std::wstring mystringW = L"mystring";
		LPBYTE bufW = LPBYTE(mystringW.c_str());
		size_t cbW = mystringW.length() * sizeof(WCHAR);
		SBinary sBinaryW = SBinary{static_cast<ULONG>(cbW), bufW};
		std::vector<BYTE> myStringWvector = std::vector<BYTE>(bufW, bufW + cbW);

		std::string mystringA = std::string("mystring");
		LPBYTE bufA = LPBYTE(mystringA.c_str());
		size_t cbA = mystringA.length();
		SBinary sBinaryA = SBinary{static_cast<ULONG>(cbA), bufA};
		std::vector<BYTE> myStringAvector = std::vector<BYTE>(bufA, bufA + cbA);

		std::vector<BYTE> vector_abcdW = std::vector<BYTE>({0x61, 0, 0x62, 0, 0x63, 0, 0x64, 0});
		std::vector<BYTE> vector_abNULLdW = std::vector<BYTE>({0x61, 0, 0x62, 0, 0x00, 0, 0x64, 0});
		std::vector<BYTE> vector_tabcrlfW = std::vector<BYTE>({0x9, 0, 0xa, 0, 0xd, 0});

		std::vector<BYTE> vector_abcdA = std::vector<BYTE>({0x61, 0x62, 0x63, 0x64});
		std::vector<BYTE> vector_abNULLdA = std::vector<BYTE>({0x61, 0x62, 0x00, 0x64});
		std::vector<BYTE> vector_tabcrlfA = std::vector<BYTE>({0x9, 0xa, 0xd});

		TEST_METHOD(Test_BinToTextStringW)
		{
			Assert::AreEqual(std::wstring(L""), strings::BinToTextStringW(nullptr, true));
			Assert::AreEqual(std::wstring(L""), strings::BinToTextStringW(nullptr, false));

			Assert::AreEqual(mystringW, strings::BinToTextStringW(&sBinaryW, false));
			Assert::AreEqual(mystringW, strings::BinToTextStringW(myStringWvector, false));
			Assert::AreEqual(std::wstring(L"abcd"), strings::BinToTextStringW(vector_abcdW, false));
			Assert::AreEqual(std::wstring(L"ab.d"), strings::BinToTextStringW(vector_abNULLdW, false));
			Assert::AreEqual(std::wstring(L"\t\n\r"), strings::BinToTextStringW(vector_tabcrlfW, true));
			Assert::AreEqual(std::wstring(L"..."), strings::BinToTextStringW(vector_tabcrlfW, false));
		}

		TEST_METHOD(Test_BinToTextString)
		{
			Assert::AreEqual(std::wstring(L""), strings::BinToTextString(nullptr, true));
			Assert::AreEqual(std::wstring(L""), strings::BinToTextString(nullptr, false));

			Assert::AreEqual(mystringW, strings::BinToTextString(&sBinaryA, false));
			auto sBinary = SBinary{static_cast<ULONG>(vector_abcdA.size()), vector_abcdA.data()};
			Assert::AreEqual(std::wstring(L"abcd"), strings::BinToTextString(&sBinary, false));
			sBinary = SBinary{static_cast<ULONG>(vector_abNULLdA.size()), vector_abNULLdA.data()};
			Assert::AreEqual(std::wstring(L"ab.d"), strings::BinToTextString(&sBinary, false));
			sBinary = SBinary{static_cast<ULONG>(vector_tabcrlfA.size()), vector_tabcrlfA.data()};
			Assert::AreEqual(std::wstring(L"\t\n\r"), strings::BinToTextString(&sBinary, true));
			sBinary = SBinary{static_cast<ULONG>(vector_tabcrlfA.size()), vector_tabcrlfA.data()};
			Assert::AreEqual(std::wstring(L"..."), strings::BinToTextString(&sBinary, false));
			Assert::AreEqual(mystringW, strings::BinToTextString(std::vector<BYTE>(bufA, bufA + cbA), true));
		}

		TEST_METHOD(Test_BinToHexString)
		{
			Assert::AreEqual(std::wstring(L"NULL"), strings::BinToHexString(std::vector<BYTE>(), false));
			Assert::AreEqual(std::wstring(L"FF"), strings::BinToHexString(std::vector<BYTE>{0xFF}, false));
			Assert::AreEqual(std::wstring(L""), strings::BinToHexString(nullptr, false));
			Assert::AreEqual(std::wstring(L"6D79737472696E67"), strings::BinToHexString(&sBinaryA, false));
			Assert::AreEqual(
				std::wstring(L"6D00790073007400720069006E006700"), strings::BinToHexString(&sBinaryW, false));
			Assert::AreEqual(
				std::wstring(L"cb: 8 lpb: 6D79737472696E67"),
				strings::BinToHexString(std::vector<BYTE>(bufA, bufA + cbA), true));
			Assert::AreEqual(std::wstring(L"6100620063006400"), strings::BinToHexString(vector_abcdW, false));
			Assert::AreEqual(std::wstring(L"6100620000006400"), strings::BinToHexString(vector_abNULLdW, false));
			Assert::AreEqual(std::wstring(L"cb: 6 lpb: 09000A000D00"), strings::BinToHexString(vector_tabcrlfW, true));
			Assert::AreEqual(std::wstring(L"09000A000D00"), strings::BinToHexString(vector_tabcrlfW, false));
		}

		TEST_METHOD(Test_HexStringToBin)
		{
			Assert::AreEqual(std::vector<BYTE>{}, strings::HexStringToBin(L"12345"));
			Assert::AreEqual(std::vector<BYTE>{}, strings::HexStringToBin(L"12345", 3));
			Assert::AreEqual(std::vector<BYTE>{0x12}, strings::HexStringToBin(L"12345", 1));
			Assert::AreEqual(std::vector<BYTE>{}, strings::HexStringToBin(L"12WZ"));
			Assert::AreEqual(vector_abcdW, strings::HexStringToBin(L"6100620063006400"));
			Assert::AreEqual(vector_abcdW, strings::HexStringToBin(L"0x6100620063006400"));
			Assert::AreEqual(vector_abcdW, strings::HexStringToBin(L"0X6100620063006400"));
			Assert::AreEqual(vector_abcdW, strings::HexStringToBin(L"x6100620063006400"));
			Assert::AreEqual(vector_abcdW, strings::HexStringToBin(L"X6100620063006400"));
			Assert::AreEqual(vector_abNULLdW, strings::HexStringToBin(L"6100620000006400"));
			Assert::AreEqual(vector_tabcrlfW, strings::HexStringToBin(L"09000A000D00"));
			Assert::AreEqual(myStringWvector, strings::HexStringToBin(L"6D00790073007400720069006E006700"));
		}

		void ByteVectorToLPBYTETest(const std::vector<BYTE>& bin) const
		{
			const auto bytes = strings::ByteVectorToLPBYTE(bin);
			for (size_t i = 0; i < bin.size(); i++)
			{
				Assert::AreEqual(bytes[i], bin[i]);
			}

			delete[] bytes;
		}

		TEST_METHOD(Test_ByteVectorToLPBYTE)
		{
			ByteVectorToLPBYTETest(std::vector<BYTE>{});
			ByteVectorToLPBYTETest(vector_abcdW);
			ByteVectorToLPBYTETest(vector_abNULLdW);
			ByteVectorToLPBYTETest(vector_tabcrlfW);

			ByteVectorToLPBYTETest(vector_abcdA);
			ByteVectorToLPBYTETest(vector_abNULLdA);
			ByteVectorToLPBYTETest(vector_tabcrlfA);
		}

		std::wstring fullstring = L"this is a string  yes";
		std::vector<std::wstring> words = {L"this", L"is", L"a", L"string", L"", L"yes"};

		TEST_METHOD(Test_split)
		{
			auto splitWords = strings::split(fullstring, L' ');

			size_t i = 0;
			for (i = 0; i < splitWords.size(); i++)
			{
				Assert::IsTrue(i < words.size());
				Assert::AreEqual(words[i], splitWords[i]);
			}

			Assert::IsTrue(i == words.size());
		}

		TEST_METHOD(Test_join) { Assert::AreEqual(fullstring, strings::join(words, L' ')); }

		TEST_METHOD(Test_currency)
		{
			Assert::AreEqual(std::wstring(L"0.0000"), strings::CurrencyToString(CURRENCY({0, 0})));
			Assert::AreEqual(std::wstring(L"858993.4593"), strings::CurrencyToString(CURRENCY({1, 2})));
		}

		TEST_METHOD(Test_base64)
		{
			Assert::AreEqual(
				std::wstring(L"bQB5AHMAdAByAGkAbgBnAA=="),
				strings::Base64Encode(myStringWvector.size(), myStringWvector.data()));
			auto ab = std::vector<byte>{0x61, 0x62};
			Assert::AreEqual(std::wstring(L"YWI="), strings::Base64Encode(ab.size(), ab.data()));
			Assert::AreEqual(myStringWvector, strings::Base64Decode(std::wstring(L"bQB5AHMAdAByAGkAbgBnAA==")));
			Assert::AreEqual(std::vector<byte>{}, strings::Base64Decode(std::wstring(L"123")));
			Assert::AreEqual(std::vector<byte>{}, strings::Base64Decode(std::wstring(L"12345===")));
			Assert::AreEqual(std::vector<byte>{}, strings::Base64Decode(std::wstring(L"12345!==")));
		}

		TEST_METHOD(Test_offsets)
		{
			Assert::AreEqual(size_t(1), strings::OffsetToFilteredOffset(L" A B", 0));
			Assert::AreEqual(size_t(3), strings::OffsetToFilteredOffset(L" A B", 1));
			Assert::AreEqual(size_t(4), strings::OffsetToFilteredOffset(L" A B", 2));
			Assert::AreEqual(static_cast<size_t>(-1), strings::OffsetToFilteredOffset(L"foo", 9));
			Assert::AreEqual(static_cast<size_t>(-1), strings::OffsetToFilteredOffset(L"", 0));
			Assert::AreEqual(size_t(0), strings::OffsetToFilteredOffset(L"foo", 0));
			Assert::AreEqual(size_t(1), strings::OffsetToFilteredOffset(L"foo", 1));
			Assert::AreEqual(size_t(2), strings::OffsetToFilteredOffset(L"foo", 2));
			Assert::AreEqual(size_t(3), strings::OffsetToFilteredOffset(L"foo", 3));
			Assert::AreEqual(size_t(5), strings::OffsetToFilteredOffset(L"    foo", 1));
			Assert::AreEqual(static_cast<size_t>(-1), strings::OffsetToFilteredOffset(L"    foo", 4));
			Assert::AreEqual(size_t(5), strings::OffsetToFilteredOffset(L"f o  o", 2));
			Assert::AreEqual(size_t(27), strings::OffsetToFilteredOffset(L"\r1\n2\t3 4-5.6,7\\8/9'a{b}c`d\"e", 13));
			Assert::AreEqual(size_t(28), strings::OffsetToFilteredOffset(L"\r1\n2\t3 4-5.6,7\\8/9'a{b}c`d\"e", 14));
		}

		TEST_METHOD(Test_endsWith)
		{
			Assert::AreEqual(true, strings::endsWith(L"1234", L"234"));
			Assert::AreEqual(true, strings::endsWith(L"Test\r\n", L"\r\n"));
			Assert::AreEqual(false, strings::endsWith(L"Test", L"\r\n"));
			Assert::AreEqual(false, strings::endsWith(L"test", L"longstring"));
		}

		TEST_METHOD(Test_ensureCRLF)
		{
			Assert::AreEqual(std::wstring(L"Test\r\n"), strings::ensureCRLF(L"Test\r\n"));
			Assert::AreEqual(std::wstring(L"Test\r\n"), strings::ensureCRLF(L"Test"));
		}

		TEST_METHOD(Test_trimWhitespace)
		{
			Assert::AreEqual(std::wstring(L"test"), strings::trimWhitespace(L" test "));
			Assert::AreEqual(std::wstring(L""), strings::trimWhitespace(L" \t \n \r "));
		}

		TEST_METHOD(Test_RemoveInvalidCharacters)
		{
			Assert::AreEqual(
				std::wstring(L"a"),
				strings::RemoveInvalidCharactersW(strings::BinToTextString(std::vector<BYTE>{0x61}, true)));
			Assert::AreEqual(std::wstring(L""), strings::RemoveInvalidCharactersW(L""));
			Assert::AreEqual(std::wstring(L".test. !"), strings::RemoveInvalidCharactersW(L"\x80test\x19\x20\x21"));
			Assert::AreEqual(
				std::wstring(L".test\r\n. !"), strings::RemoveInvalidCharactersW(L"\x80test\r\n\x19\x20\x21", true));
			auto nullTerminatedStringW = std::wstring{L"test"};
			nullTerminatedStringW.push_back(L'\0');
			Assert::AreEqual(nullTerminatedStringW, strings::RemoveInvalidCharactersW(nullTerminatedStringW, true));
			Assert::AreEqual(std::wstring(L"アメリカ"), strings::RemoveInvalidCharactersW(L"アメリカ"));

			Assert::AreEqual(
				std::string("a"),
				strings::RemoveInvalidCharactersA(
					strings::wstringTostring(strings::BinToTextString(std::vector<BYTE>{0x61}, true))));
			Assert::AreEqual(std::string(""), strings::RemoveInvalidCharactersA(""));
			Assert::AreEqual(std::string(".test. !"), strings::RemoveInvalidCharactersA("\x80test\x19\x20\x21"));
			auto nullTerminatedStringA = std::string{"test"};
			nullTerminatedStringA.push_back('\0');
			Assert::AreEqual(nullTerminatedStringA, strings::RemoveInvalidCharactersA(nullTerminatedStringA, true));
		}

		TEST_METHOD(Test_FileTimeToString)
		{
			std::wstring prop;
			std::wstring alt;
			strings::FileTimeToString(FILETIME{0x085535B0, 0x01D387EC}, prop, alt);
			Assert::AreEqual(std::wstring(L"07:16:34.571 PM 1/7/2018"), prop);
			Assert::AreEqual(std::wstring(L"Low: 0x085535B0 High: 0x01D387EC"), alt);

			strings::FileTimeToString(FILETIME{0xFFFFFFFF, 0xFFFFFFFF}, prop, alt);
			Assert::AreEqual(std::wstring(L"Invalid systime"), prop);
			Assert::AreEqual(std::wstring(L"Low: 0xFFFFFFFF High: 0xFFFFFFFF"), alt);
		}
	};
} // namespace stringtest