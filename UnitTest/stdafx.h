#pragma once
#include "targetver.h"
#include "CppUnitTest.h"
#include "..\StdAfx.h"

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			template<> inline std::wstring ToString<__int64>(const __int64& t) { RETURN_WIDE_STRING(t); }
			template<> inline std::wstring ToString<vector<BYTE>>(const vector<BYTE>& t) { RETURN_WIDE_STRING(t.data()); }

			// Assert::AreEqual doesn't do full logging, so we roll our own
			void inline AreEqualEx(const wstring& expected, const wstring& actual, const wchar_t* message = nullptr, const __LineInfo* pLineInfo = nullptr)
			{
				if (expected != actual)
				{
					Logger::WriteMessage(format(L"Expected:\n%ws\n", expected.c_str()).c_str());
					Logger::WriteMessage(format(L"\nActual:\n%ws\n", actual.c_str()).c_str());
					Assert::IsTrue(false, ToString(message).c_str(), pLineInfo);
				}
			}
		}
	}
}