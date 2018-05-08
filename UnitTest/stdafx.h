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
			static const bool AssertOnFailure = true;
			static const bool LimitOutput = false;
			template<> inline std::wstring ToString<__int64>(const __int64& t) { RETURN_WIDE_STRING(t); }
			template<> inline std::wstring ToString<vector<BYTE>>(const vector<BYTE>& t) { RETURN_WIDE_STRING(t.data()); }

			// Assert::AreEqual doesn't do full logging, so we roll our own
			void inline AreEqualEx(const wstring& expected, const wstring& actual, const wchar_t* message = nullptr, const __LineInfo* pLineInfo = nullptr)
			{
				if (expected != actual)
				{
					if (message != nullptr)
					{
						Logger::WriteMessage(format(L"Test: %ws\n", message).c_str());
					}

					Logger::WriteMessage(L"Diff:\n");

					auto splitExpected = split(expected, L'\n');
					auto splitActual = split(actual, L'\n');
					auto errorCount = 0;
					for (size_t i = 0; i < splitExpected.size() && i < splitActual.size() && (errorCount < 4 || !LimitOutput); i++)
					{
						if (splitExpected[i] != splitActual[i])
						{
							errorCount++;
							Logger::WriteMessage(format(L"[%d]\n\"%ws\"\n\"%ws\"\n", i, splitExpected[i].c_str(), splitActual[i].c_str()).c_str());
						}
					}

					if (!LimitOutput)
					{
						Logger::WriteMessage(L"\n");
						Logger::WriteMessage(format(L"Expected:\n\"%ws\"\n\n", expected.c_str()).c_str());
						Logger::WriteMessage(format(L"Actual:\n\"%ws\"", actual.c_str()).c_str());
					}

					if (AssertOnFailure)
					{
						Assert::Fail(ToString(message).c_str(), pLineInfo);
					}
				}
			}
		}
	}
}