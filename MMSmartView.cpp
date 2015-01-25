#include "stdafx.h"

#include "MrMAPI.h"
#include "MAPIFunctions.h"
#include "String.h"
#include <io.h>
#include "SmartView\SmartView.h"

void DoSmartView(_In_ MYOPTIONS ProgOpts)
{
	// Ignore the reg key that disables smart view parsing
	RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD = true;

	__ParsingTypeEnum ulStructType = IDS_STNOPARSING;

	if (ProgOpts.ulSVParser && ProgOpts.ulSVParser < ulSmartViewParserTypeArray)
	{
		ulStructType = (__ParsingTypeEnum)SmartViewParserTypeArray[ProgOpts.ulSVParser].ulValue;
	}

	if (ulStructType)
	{
		FILE* fIn = NULL;
		FILE* fOut = NULL;
		fIn = _wfopen(ProgOpts.lpszInput, L"rb");
		if (!fIn) printf("Cannot open input file %ws\n", ProgOpts.lpszInput);
		if (ProgOpts.lpszOutput)
		{
			fOut = _wfopen(ProgOpts.lpszOutput, L"wb");
			if (!fOut) printf("Cannot open output file %ws\n", ProgOpts.lpszOutput);
		}

		if (fIn && (!ProgOpts.lpszOutput || fOut))
		{
			int iDesc = _fileno(fIn);
			long iLength = _filelength(iDesc);

			LPBYTE lpbIn = new BYTE[iLength + 1]; // +1 for NULL
			if (lpbIn)
			{
				memset(lpbIn, 0, sizeof(BYTE)*(iLength + 1));
				fread(lpbIn, sizeof(BYTE), iLength, fIn);
				SBinary Bin = { 0 };
				if (ProgOpts.ulOptions & OPT_BINARYFILE)
				{
					Bin.cb = iLength;
					Bin.lpb = lpbIn;
				}
				else
				{
					LPTSTR szIn = NULL;
#ifdef UNICODE
					EC_H(AnsiToUnicode((LPCSTR) lpbIn,&szIn));
#else
					szIn = (LPSTR)lpbIn;
#endif
					if (MyBinFromHex(
						(LPCTSTR)szIn,
						NULL,
						&Bin.cb))
					{
						Bin.lpb = new BYTE[Bin.cb];
						if (Bin.lpb)
						{
							(void)MyBinFromHex(
								(LPCTSTR)szIn,
								Bin.lpb,
								&Bin.cb);
						}
					}
#ifdef UNICODE
					delete[] szIn;
#endif
				}

				wstring szString = InterpretBinaryAsString(Bin, ulStructType, NULL, NULL);
				if (!szString.empty())
				{
					if (fOut)
					{
						// Without this split, the ANSI build writes out UNICODE files
#ifdef UNICODE
						fputws(szString.c_str(),fOut);
#else
						LPSTR szStringA = NULL;
						(void)UnicodeToAnsi(szString.c_str(), &szStringA);
						fputs(szStringA, fOut);
						delete[] szStringA;
#endif
					}
					else
					{
						_tprintf(_T("%ws\n"), szString.c_str());
					}
				}

				if (!(ProgOpts.ulOptions & OPT_BINARYFILE)) delete[] Bin.lpb;
				delete[] lpbIn;
			}
		}

		if (fOut) fclose(fOut);
		if (fIn) fclose(fIn);
	}
}