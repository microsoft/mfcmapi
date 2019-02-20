#include <StdAfx.h>
#include <MrMapi/MMErr.h>
#include <MrMapi/mmcli.h>
#include <core/utility/strings.h>

namespace error
{
	extern ERROR_ARRAY_ENTRY g_ErrorArray[];
	extern ULONG g_ulErrorArray;
} // namespace error

void PrintErrFromNum(_In_ ULONG ulError)
{
	printf("0x%08lX = %ws\n", ulError, error::ErrorNameFromErrorCode(ulError).c_str());
}

void PrintErrFromName(_In_ const std::wstring& lpszError)
{
	const auto szErr = lpszError.c_str();

	for (ULONG i = 0; i < error::g_ulErrorArray; i++)
	{
		if (0 == lstrcmpiW(szErr, error::g_ErrorArray[i].lpszName))
		{
			printf("0x%08lX = %ws\n", error::g_ErrorArray[i].ulErrorName, szErr);
		}
	}
}

void PrintErrFromPartialName(_In_ const std::wstring& lpszError)
{
	if (!lpszError.empty())
		printf("Searching for \"%ws\"\n", lpszError.c_str());
	else
		printf("Searching for all errors\n");

	ULONG ulNumMatches = 0;

	for (ULONG ulCur = 0; ulCur < error::g_ulErrorArray; ulCur++)
	{
		if (lpszError.empty() || nullptr != StrStrIW(error::g_ErrorArray[ulCur].lpszName, lpszError.c_str()))
		{
			printf("0x%08lX = %ws\n", error::g_ErrorArray[ulCur].ulErrorName, error::g_ErrorArray[ulCur].lpszName);
			ulNumMatches++;
		}
	}

	printf("Found %lu matches.\n", ulNumMatches);
}

void DoErrorParse(_In_ cli::OPTIONS ProgOpts)
{
	auto lpszErr = ProgOpts.lpszUnswitchedOption;
	const auto ulErrNum = strings::wstringToUlong(lpszErr, ProgOpts.optionFlags & cli::OPT_DODECIMAL ? 10 : 16);

	if (ulErrNum)
	{
		PrintErrFromNum(ulErrNum);
	}
	else
	{
		if (ProgOpts.optionFlags & cli::OPT_DOPARTIALSEARCH || lpszErr.empty())
		{
			PrintErrFromPartialName(lpszErr);
		}
		else
		{
			PrintErrFromName(lpszErr);
		}
	}
}