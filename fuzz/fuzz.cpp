#include <StdAfx.h>
#include <core/smartview/SmartView.h>
#include <core/addin/mfcmapi.h>
#include <core/addin/addin.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <array>
#include <chrono>

#ifdef FUZZ
std::once_flag _initFlag;
void EnsureInit()
{
	addin::MergeAddInArrays();
	registry::doSmartView = true;
	registry::useGetPropList = true;
	registry::parseNamedProps = true;
	registry::cacheNamedProps = true;
	strings::setTestInstance(GetModuleHandleW(L"mfcmapi.exe"));
}

void test(const SBinary hex)
{
	static auto testnum = 1LL;
	const auto batchsize = 1000ll;
	const bool doLog = testnum++ % batchsize == 0;

	// we're gonna track run times for each parser type in an array
	// We add the time of the current test to it's bucket, then every time doLog is true, we print the average time for that parser type
	static std::array<std::chrono::duration<double, std::milli>, static_cast<size_t>(parserType::END)> runtimes = {};
	for (const auto parser : SmartViewParserTypeArray)
	{
		if (parser.type == parserType::NOPARSING) continue;
		//wprintf(L"Testing %ws\r\n", addin::AddInStructTypeToString(parser.type).c_str());
		const auto start = std::chrono::high_resolution_clock::now();
		(void) smartview::InterpretBinary(hex, parser.type, nullptr);
		const auto end = std::chrono::high_resolution_clock::now();
		runtimes[static_cast<size_t>(parser.type)] += end - start;
	}

	if (doLog)
	{
		wprintf(L"Test %lld\r\n", testnum);
		wprintf(L"%-40s %-15s %-15s\n", L"Parser Type", L"Total Time (ms)", L"Time per Test (ms)");
		for (const auto parser : SmartViewParserTypeArray)
		{
			if (parser.type == parserType::NOPARSING) continue;
			wprintf(
				L"%-40ws %-15f %-15f\n",
				addin::AddInStructTypeToString(parser.type).c_str(),
				runtimes[static_cast<size_t>(parser.type)].count(),
				runtimes[static_cast<size_t>(parser.type)].count() / testnum);
		}
	}
}

#ifdef __cplusplus
#define FUZZ_EXPORT extern "C" __declspec(dllexport)
#else
#define FUZZ_EXPORT __declspec(dllexport)
#endif
FUZZ_EXPORT int __cdecl LLVMFuzzerTestOneInput(const uint8_t* data, size_t size)
{
	std::call_once(_initFlag, EnsureInit);
	const SBinary input = {static_cast<ULONG>(size), (LPBYTE) (data)};
	//wprintf(L"Fuzzing: %ws\r\n", strings::BinToHexString(&input, true).c_str());
	test(input);
	return 0;
}
#endif // FUZZ