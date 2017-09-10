#include "stdafx.h"
#include "MrMAPI.h"
#include <Interpret/String.h>
#include <io.h>
#include <Interpret/SmartView/SmartView.h>

void DoSmartView(_In_ MYOPTIONS ProgOpts)
{
	// Ignore the reg key that disables smart view parsing
	RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD = true;

	auto ulStructType = IDS_STNOPARSING;

	if (ProgOpts.ulSVParser && ProgOpts.ulSVParser < SmartViewParserTypeArray.size())
	{
		ulStructType = static_cast<__ParsingTypeEnum>(SmartViewParserTypeArray[ProgOpts.ulSVParser].ulValue);
	}

	if (ulStructType)
	{
		FILE* fOut = nullptr;
		auto fIn = MyOpenFile(ProgOpts.lpszInput.c_str(), L"rb");
		if (!fIn) printf("Cannot open input file %ws\n", ProgOpts.lpszInput.c_str());
		if (!ProgOpts.lpszOutput.empty())
		{
			fOut = MyOpenFile(ProgOpts.lpszOutput.c_str(), L"wb");
			if (!fOut) printf("Cannot open output file %ws\n", ProgOpts.lpszOutput.c_str());
		}

		if (fIn && (ProgOpts.lpszOutput.empty() || fOut))
		{
			auto iDesc = _fileno(fIn);
			auto iLength = _filelength(iDesc);

			auto lpbIn = new BYTE[iLength + 1]; // +1 for NULL
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
					bin = HexStringToBin(LPCSTRToWstring(reinterpret_cast<LPCSTR>(lpbIn)));
					Bin.cb = static_cast<ULONG>(bin.size());
					Bin.lpb = bin.data();
				}

				auto szString = InterpretBinaryAsString(Bin, ulStructType, nullptr);
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