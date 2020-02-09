#include <StdAfx.h>
#include <core/smartview/SmartView.h>
#include <core/utility/strings.h>
#include <MrMapi/mmcli.h>
#include <io.h>
#include <core/addin/addin.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>

void DoSmartView()
{
	// Ignore the reg key that disables smart view parsing
	registry::doSmartView = true;

	auto ulStructType = IDS_STNOPARSING;
	const auto ulSVParser = cli::switchParser.atULONG(0);
	if (ulSVParser && ulSVParser < SmartViewParserTypeArray.size())
	{
		ulStructType = static_cast<parserType>(SmartViewParserTypeArray[ulSVParser].ulValue);
	}

	if (ulStructType)
	{
		FILE* fOut = nullptr;
		const auto input = cli::switchInput[0];
		const auto output = cli::switchOutput[0];
		const auto fIn = output::MyOpenFileMode(input, L"rb");
		if (!fIn) printf("Cannot open input file %ws\n", input.c_str());
		if (!output.empty())
		{
			fOut = output::MyOpenFileMode(output, L"wb");
			if (!fOut) printf("Cannot open output file %ws\n", output.c_str());
		}

		if (fIn && (output.empty() || fOut))
		{
			const auto iDesc = _fileno(fIn);
			const auto iLength = _filelength(iDesc);

			auto inBytes = std::vector<BYTE>(iLength + 1); // +1 for NULL
			fread(inBytes.data(), sizeof(BYTE), iLength, fIn);
			auto sBin = SBinary{};
			auto bin = std::vector<BYTE>{};
			if (cli::switchBinary.isSet())
			{
				sBin.cb = iLength;
				sBin.lpb = inBytes.data();
			}
			else
			{
				bin = strings::HexStringToBin(strings::LPCSTRToWstring(reinterpret_cast<LPCSTR>(inBytes.data())));
				sBin.cb = static_cast<ULONG>(bin.size());
				sBin.lpb = bin.data();
			}

			auto szString = smartview::InterpretBinaryAsString(sBin, ulStructType, nullptr);
			if (!szString.empty())
			{
				if (fOut)
				{
					output::Output(output::dbgLevel::NoDebug, fOut, false, szString);
				}
				else
				{
					wprintf(L"%ws\n", strings::StripCarriage(szString).c_str());
				}
			}
		}

		if (fOut) fclose(fOut);
		if (fIn) fclose(fIn);
	}
}
