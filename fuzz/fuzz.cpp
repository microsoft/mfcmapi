#include <StdAfx.h>
#include <core/smartview/SmartView.h>
#include <core/addin/mfcmapi.h>
#include <core/addin/addin.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>

#ifdef FUZZ
std::once_flag _initFlag;
void EnsureInit()
{
	addin::MergeAddInArrays();
	registry::doSmartView = true;
	registry::useGetPropList = true;
	registry::parseNamedProps = true;
	registry::cacheNamedProps = true;
	strings::setTestInstance(GetModuleHandleW(L"fuzz.exe"));
}

void test(const SBinary hex)
{
	for (const auto parser : SmartViewParserTypeArray)
	{
		if (parser.type == parserType::NOPARSING) continue;
		//wprintf(L"Testing %ws\r\n", addin::AddInStructTypeToString(parser.type).c_str());
		(void) smartview::InterpretBinary(hex, parser.type, nullptr);
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