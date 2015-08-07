#include "stdafx.h"
#include "..\stdafx.h"

#include "MrMAPI.h"
#include "MMErr.h"
#include <shlwapi.h>
#include <io.h>

void PrintErrFromNum(_In_ ULONG ulError)
{
	LPWSTR szErr = ErrorNameFromErrorCode(ulError);
	printf("0x%08X = %ws\n",ulError,szErr);
} // PrintErrFromNum

void PrintErrFromName(_In_z_ LPCWSTR lpszError)
{
	ULONG i = 0;

	for (i = 0;i < g_ulErrorArray;i++)
	{
		if (0 == lstrcmpiW(lpszError,g_ErrorArray[i].lpszName))
		{
			printf("0x%08X = %ws\n",g_ErrorArray[i].ulErrorName,lpszError);
		}
	}
} // PrintErrFromName

void PrintErrFromPartialName(_In_opt_z_ LPCWSTR lpszError)
{
	if (lpszError) printf("Searching for \"%ws\"\n",lpszError);
	else printf("Searching for all errors\n");

	ULONG ulCur = 0;
	ULONG ulNumMatches = 0;

	for (ulCur = 0 ; ulCur < g_ulErrorArray ; ulCur++)
	{
		if (!lpszError || 0 != StrStrIW(g_ErrorArray[ulCur].lpszName,lpszError))
		{
			printf("0x%08X = %ws\n",g_ErrorArray[ulCur].ulErrorName,g_ErrorArray[ulCur].lpszName);
			ulNumMatches++;
		}
	}
	printf("Found %u matches.\n",ulNumMatches);
} // PrintErrFromPartialName

void DoErrorParse(_In_ MYOPTIONS ProgOpts)
{
	ULONG ulErrNum = NULL;
	LPWSTR lpszErr = ProgOpts.lpszUnswitchedOption;

	if (lpszErr)
	{
		ULONG ulArg = NULL;
		LPWSTR szEndPtr = NULL;
		ulArg = wcstoul(lpszErr,&szEndPtr,(ProgOpts.ulOptions & OPT_DODECIMAL)?10:16);

		// if szEndPtr is pointing to something other than NULL, this must be a string
		if (!szEndPtr || *szEndPtr)
		{
			ulArg = NULL;
		}

		ulErrNum = ulArg;
	}

	if (ulErrNum)
	{
		PrintErrFromNum(ulErrNum);
	}
	else
	{
		if ((ProgOpts.ulOptions & OPT_DOPARTIALSEARCH) || !lpszErr)
		{
			PrintErrFromPartialName(lpszErr);
		}
		else
		{
			PrintErrFromName(lpszErr);
		}
	}
} // DoErrorParse