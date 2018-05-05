#pragma once
#include "..\StdAfx.h"
#include "targetver.h"
#include "CppUnitTest.h"
#include "Interpret\String.h"

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
					Logger::WriteMessage(L"Diff:\n");

					auto splitExpected = split(expected, L'\n');
					auto splitActual = split(actual, L'\n');
					for (size_t i = 0; i < splitExpected.size() && i < splitActual.size(); i++)
					{
						if (splitExpected[i] != splitActual[i])
						{
							Logger::WriteMessage(format(L"[%d]\n\"%ws\"\n\"%ws\"\n", i, splitExpected[i].c_str(), splitActual[i].c_str()).c_str());
						}
					}

					Logger::WriteMessage(L"\n");
					Logger::WriteMessage(format(L"Expected:\n\"%ws\"\n\n", expected.c_str()).c_str());
					Logger::WriteMessage(format(L"Actual:\n\"%ws\"", actual.c_str()).c_str());

					Assert::IsTrue(false, ToString(message).c_str(), pLineInfo);
				}
			}
		}
	}
}