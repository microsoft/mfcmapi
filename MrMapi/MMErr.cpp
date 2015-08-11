#include "stdafx.h"
#include "..\stdafx.h"

#include "MrMAPI.h"
#include "MMErr.h"
#include <shlwapi.h>
#include <io.h>
#include "..\String.h"

void PrintErrFromNum(_In_ ULONG ulError)
{
	printf("0x%08X = %ws\n", ulError, ErrorNameFromErrorCode(ulError));
}

void PrintErrFromName(_In_ wstring lpszError)
{
	LPCWSTR szErr = lpszError.c_str();
	ULONG i = 0;

	for (i = 0; i < g_ulErrorArray; i++)
	{
		if (0 == lstrcmpiW(szErr, g_ErrorArray[i].lpszName))
		{
			printf("0x%08X = %ws\n", g_ErrorArray[i].ulErrorName, szErr);
		}
	}
}

void PrintErrFromPartialName(_In_ wstring lpszError)
{
	if (!lpszError.empty()) printf("Searching for \"%ws\"\n", lpszError.c_str());
	else printf("Searching for all errors\n");

	ULONG ulCur = 0;
	ULONG ulNumMatches = 0;

	for (ulCur = 0; ulCur < g_ulErrorArray; ulCur++)
	{
		if (lpszError.empty() || 0 != StrStrIW(g_ErrorArray[ulCur].lpszName, lpszError.c_str()))
		{
			printf("0x%08X = %ws\n", g_ErrorArray[ulCur].ulErrorName, g_ErrorArray[ulCur].lpszName);
			ulNumMatches++;
		}
	}

	printf("Found %u matches.\n", ulNumMatches);
}

void DoErrorParse(_In_ MYOPTIONS ProgOpts)
{
	wstring lpszErr = ProgOpts.lpszUnswitchedOption;
	ULONG ulErrNum = wstringToUlong(lpszErr, (ProgOpts.ulOptions & OPT_DODECIMAL) ? 10 : 16);

	if (ulErrNum)
	{
		PrintErrFromNum(ulErrNum);
	}
	else
	{
		if ((ProgOpts.ulOptions & OPT_DOPARTIALSEARCH) || lpszErr.empty())
		{
			PrintErrFromPartialName(lpszErr);
		}
		else
		{
			PrintErrFromName(lpszErr);
		}
	}
}