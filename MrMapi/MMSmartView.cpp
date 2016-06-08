#include "stdafx.h"
#include "..\stdafx.h"
#include "MrMAPI.h"
#include "..\MAPIFunctions.h"
#include "..\String.h"
#include <io.h>
#include "..\SmartView\SmartView.h"

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
		fIn = _wfopen(ProgOpts.lpszInput.c_str(), L"rb");
		if (!fIn) printf("Cannot open input file %ws\n", ProgOpts.lpszInput.c_str());
		if (!ProgOpts.lpszOutput.empty())
		{
			fOut = _wfopen(ProgOpts.lpszOutput.c_str(), L"wb");
			if (!fOut) printf("Cannot open output file %ws\n", ProgOpts.lpszOutput.c_str());
		}

		if (fIn && (ProgOpts.lpszOutput.empty() || fOut))
		{
			int iDesc = _fileno(fIn);
			long iLength = _filelength(iDesc);

			LPBYTE lpbIn = new BYTE[iLength + 1]; // +1 for NULL
			if (lpbIn)
			{
				memset(lpbIn, 0, sizeof(BYTE)*(iLength + 1));
				fread(lpbIn, sizeof(BYTE), iLength, fIn);
				SBinary Bin = { 0 };
				vector<BYTE> bin;
				if (ProgOpts.ulOptions & OPT_BINARYFILE)
				{
					Bin.cb = iLength;
					Bin.lpb = lpbIn;
				}
				else
				{
					wstring str = LPCSTRToWstring((LPCSTR)lpbIn);
					bin = HexStringToBin(str);
					Bin.cb = bin.size();
					Bin.lpb = bin.data();
				}

				wstring szString = InterpretBinaryAsString(Bin, ulStructType, NULL);
				if (!szString.empty())
				{
					if (fOut)
					{
						Output(DBGNoDebug, fOut, false, szString);
					}
					else
					{
						wprintf(L"%ws\n", StripCarriage(szString).c_str());
					}
				}

				delete[] lpbIn;
			}
		}

		if (fOut) fclose(fOut);
		if (fIn) fclose(fIn);
	}
}