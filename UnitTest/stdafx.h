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
		}
	}
}