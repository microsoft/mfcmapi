#include <StdAfx.h>
#include <MrMapi/MMSmartView.h>
#include <core/smartview/SmartView.h>
#include <core/utility/strings.h>
#include <MrMapi/mmcli.h>
#include <io.h>
#include <core/addin/addin.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/error.h>
#include <limits>
#include <memory>
#include <stdexcept>

namespace smartviewfile
{
	namespace
	{
		using File = std::unique_ptr<FILE, decltype(&output::CloseFile)>;

		HRESULT Failure(const wchar_t* message, HRESULT result)
		{
			wprintf(L"SmartView: %ws (0x%08lX).\n", message, result);
			return WC_H(result);
		}

		HRESULT PrepareOutput(FILE* input, FILE* outputFile)
		{
			BY_HANDLE_FILE_INFORMATION inputInfo{};
			BY_HANDLE_FILE_INFORMATION outputInfo{};
			if (!GetFileInformationByHandle(reinterpret_cast<HANDLE>(_get_osfhandle(_fileno(input))), &inputInfo) ||
				!GetFileInformationByHandle(reinterpret_cast<HANDLE>(_get_osfhandle(_fileno(outputFile))), &outputInfo))
				return Failure(L"Cannot identify input and output files", HRESULT_FROM_WIN32(GetLastError()));

			if (inputInfo.dwVolumeSerialNumber == outputInfo.dwVolumeSerialNumber &&
				inputInfo.nFileIndexHigh == outputInfo.nFileIndexHigh &&
				inputInfo.nFileIndexLow == outputInfo.nFileIndexLow)
				return Failure(
					L"Input and output must be different files", HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER));

			if (_chsize_s(_fileno(outputFile), 0) || _fseeki64(outputFile, 0, SEEK_SET))
				return Failure(L"Cannot prepare output file", HRESULT_FROM_WIN32(ERROR_WRITE_FAULT));
			return S_OK;
		}
	} // namespace

	const InputOperations& DefaultInputOperations()
	{
		static const InputOperations operations{
			[](FILE* file) { return _filelengthi64(_fileno(file)); },
			[](BYTE* bytes, size_t size, FILE* file) { return fread(bytes, 1, size, file); },
			[](FILE* file) { return ferror(file) != 0; },
			[](SBinary bytes, parserType parser) {
				return smartview::InterpretBinary(bytes, parser, nullptr)->toString();
			}};
		return operations;
	}

	HRESULT ValidateInputSize(__int64 length, size_t bufferMaxSize, size_t& size)
	{
		size = 0;
		if (length < 0) return WC_H(HRESULT_FROM_WIN32(ERROR_FILE_INVALID));
		const auto value = static_cast<unsigned __int64>(length);
		if (value > (std::numeric_limits<ULONG>::max)() || value >= (std::numeric_limits<size_t>::max)() ||
			value >= bufferMaxSize)
			return WC_H(HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE));
		size = static_cast<size_t>(value);
		return S_OK;
	}

	HRESULT ParseFile(
		const std::wstring& input,
		const std::wstring& output,
		parserType parser,
		bool binary,
		const InputOperations& operations)
	{
		try
		{
			File fIn{output::MyOpenFileMode(input, L"rb"), output::CloseFile};
			if (!fIn) return Failure(L"Cannot open input file", HRESULT_FROM_WIN32(ERROR_OPEN_FAILED));

			std::vector<BYTE> inBytes;
			size_t length{};
			const auto result = ValidateInputSize(operations.size(fIn.get()), inBytes.max_size(), length);
			if (FAILED(result)) return Failure(L"Cannot determine a supported input file size", result);

			inBytes.resize(length + 1);
			const auto read = operations.read(inBytes.data(), length, fIn.get());
			if (read != length || operations.error(fIn.get()))
				return Failure(L"Input file could not be read completely", HRESULT_FROM_WIN32(ERROR_READ_FAULT));
			inBytes[read] = 0;
			auto sBin = SBinary{};
			auto bin = std::vector<BYTE>{};
			if (binary)
			{
				sBin.cb = static_cast<ULONG>(read);
				sBin.lpb = inBytes.data();
			}
			else
			{
				bin = strings::HexStringToBin(strings::LPCSTRToWstring(reinterpret_cast<LPCSTR>(inBytes.data())));
				if (bin.size() > (std::numeric_limits<ULONG>::max)())
					return Failure(L"Decoded input is too large", HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE));
				sBin.cb = static_cast<ULONG>(bin.size());
				sBin.lpb = bin.data();
			}

			const auto szString = operations.parse(sBin, parser);
			if (output.empty())
			{
				if (!szString.empty() && wprintf(L"%ws\n", strings::StripCarriage(szString).c_str()) < 0)
					return Failure(L"Cannot write parsed output", HRESULT_FROM_WIN32(ERROR_WRITE_FAULT));
			}
			else
			{
				// Append mode obtains a handle without truncating an existing input alias.
				File fOut{output::MyOpenFileMode(output, L"ab"), output::CloseFile};
				if (!fOut) return Failure(L"Cannot open output file", HRESULT_FROM_WIN32(ERROR_OPEN_FAILED));
				const auto prepared = PrepareOutput(fIn.get(), fOut.get());
				if (FAILED(prepared)) return prepared;
				if (!szString.empty()) output::Output(output::dbgLevel::NoDebug, fOut.get(), false, szString);
				if (ferror(fOut.get()) || fflush(fOut.get()))
				{
					return Failure(L"Cannot write parsed output", HRESULT_FROM_WIN32(ERROR_WRITE_FAULT));
				}
				if (fclose(fOut.release()))
					return Failure(L"Cannot close output file", HRESULT_FROM_WIN32(ERROR_WRITE_FAULT));
			}
			return S_OK;
		} catch (const std::bad_alloc&)
		{
			return Failure(L"Not enough memory to process input", E_OUTOFMEMORY);
		} catch (const std::length_error&)
		{
			return Failure(L"Input exceeds the supported buffer size", HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE));
		}
	}
} // namespace smartviewfile

HRESULT DoSmartView()
{
	// Ignore the reg key that disables smart view parsing
	registry::doSmartView = true;

	const auto parser = cli::switchParser.atULONG(0);
	if (!parser || parser >= SmartViewParserTypeArray.size())
	{
		wprintf(L"SmartView: Invalid parser type.\n");
		return WC_H(E_INVALIDARG);
	}
	return smartviewfile::ParseFile(
		cli::switchInput[0],
		cli::switchOutput[0],
		static_cast<parserType>(SmartViewParserTypeArray[parser].type),
		cli::switchBinary.isSet());
}
