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

void test(std::vector<BYTE> hex)
{
	for (const auto parser : SmartViewParserTypeArray)
	{
		if (parser.type == parserType::NOPARSING) continue;
		wprintf(L"Testing %ws\r\n", addin::AddInStructTypeToString(parser.type).c_str());
		(void) smartview::InterpretBinary({static_cast<ULONG>(hex.size()), hex.data()}, parser.type, nullptr);
	}
}

std::wstring LoadDataToString(const uint8_t* Data, size_t Size)
{
	const auto cb = Size;
	const LPVOID bytes = (LPVOID) Data;
	const auto data = static_cast<const BYTE*>(bytes);

	// UTF 16 LE
	// In Notepad++, this is UCS-2 LE BOM encoding
	// WARNING: Editing files in Visual Studio Code can alter this encoding
	if (cb >= 2 && data[0] == 0xff && data[1] == 0xfe)
	{
		// Skip the byte order mark
		const auto wstr = static_cast<const wchar_t*>(bytes);
		const auto cch = cb / sizeof(wchar_t);
		return std::wstring(wstr + 1, cch - 1);
	}

	const auto str = std::string(static_cast<const char*>(bytes), cb);
	return strings::stringTowstring(str);
}

#ifdef __cplusplus
#define FUZZ_EXPORT extern "C" __declspec(dllexport)
#else
#define FUZZ_EXPORT __declspec(dllexport)
#endif
FUZZ_EXPORT int __cdecl LLVMFuzzerTestOneInput(const uint8_t* data, size_t size)
{
	std::call_once(_initFlag, EnsureInit);
	// convert data to vector byte

	const auto inputVector = std::vector<BYTE>(data, data + size);
	const auto input = LoadDataToString(data, size);
	if (input.empty())
	{
		// Print hex encoding of input so we can see what was wrong with it
		wprintf(L"Invalid input: %ws\r\n", strings::BinToHexString(inputVector, true).c_str());
		return -1; // ignore invalid hex strings
	}

	auto hex = strings::HexStringToBin(input);
	if (hex.empty())
	{
		wprintf(L"Invalid hex: %ws\r\n", input.c_str());
		return -1; // ignore invalid hex strings
	}

	wprintf(L"Fuzzing: %ws\r\n", input.c_str());
	test(hex);
	return 0;
}
#endif // FUZZ