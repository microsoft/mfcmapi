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
			template<> inline std::wstring ToString<std::vector<BYTE>>(const std::vector<BYTE>& t) { RETURN_WIDE_STRING(t.data()); }
		}
	}
}