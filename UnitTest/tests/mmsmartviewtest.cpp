#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <MrMapi/MMSmartView.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>
#include <core/smartview/SmartView.h>
#include <core/utility/output.h>
#include <core/utility/strings.h>
#include <filesystem>
#include <io.h>
#include <limits>
#include <winioctl.h>

namespace mmsmartviewtest
{
	using File = std::unique_ptr<FILE, decltype(&output::CloseFile)>;

	struct CloseHandle
	{
		void operator()(void* value) const { ::CloseHandle(value); }
	};
	using Handle = std::unique_ptr<void, CloseHandle>;

	TEST_CLASS(mmsmartviewtest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD_INITIALIZE(createFiles)
		{
			m_input = TemporaryFile();
			m_output = TemporaryFile();
			m_alias.clear();
			m_bytes.assign(24, 0);
			m_bytes[0] = 1;
			Write(m_input, m_bytes);
			Write(m_output, m_sentinel);
			m_reads = 0;
			m_parses = 0;
			m_parsed.clear();
			m_operations = smartviewfile::DefaultInputOperations();
			m_operations.read = [this](BYTE* bytes, size_t size, FILE* file) {
				++m_reads;
				return smartviewfile::DefaultInputOperations().read(bytes, size, file);
			};
			m_operations.parse = [this](SBinary bytes, parserType parser) {
				++m_parses;
				m_parsed.clear();
				if (bytes.cb) m_parsed.assign(bytes.lpb, bytes.lpb + bytes.cb);
				return smartviewfile::DefaultInputOperations().parse(bytes, parser);
			};
		}

		TEST_METHOD_CLEANUP(removeFiles)
		{
			for (const auto& file : {m_alias, m_input, m_output})
			{
				if (file.empty()) continue;
				Assert::IsTrue(DeleteFileW(file.c_str()) != FALSE || GetLastError() == ERROR_FILE_NOT_FOUND);
			}
		}

		TEST_METHOD(BinaryInputRetainsParsedOutputAndExactByteCount)
		{
			Assert::AreEqual(S_OK, Parse(true));
			Assert::IsTrue(m_parsed == m_bytes);
			Assert::AreEqual(1UL, m_reads);
			Assert::AreEqual(1UL, m_parses);
			CheckOutput(m_bytes);
		}

		TEST_METHOD(HexInputWithoutTerminatorRetainsParsedOutput)
		{
			const auto hex = strings::BinToHexString(m_bytes, false);
			const auto text = strings::wstringTostring(hex);
			Write(m_input, std::vector<BYTE>(text.begin(), text.end()));
			Assert::AreEqual(S_OK, Parse(false));
			Assert::IsTrue(m_parsed == m_bytes);
			Assert::AreEqual(1UL, m_parses);
			CheckOutput(m_bytes);
		}

		TEST_METHOD(EmptyBinaryAndHexInputsAreParsedAsEmptyBlobs)
		{
			for (const auto binary : {false, true})
			{
				Write(m_input, {});
				Assert::AreEqual(S_OK, Parse(binary));
				Assert::IsTrue(m_parsed.empty());
				CheckOutput({});
			}
			Assert::AreEqual(2UL, m_parses);
		}

		TEST_METHOD(MissingInputPreservesExistingOutput)
		{
			Assert::IsTrue(DeleteFileW(m_input.c_str()) != FALSE);
			CheckFailure(Parse(true));
			Assert::AreEqual(0UL, m_reads);
		}

		TEST_METHOD(UnreadableInputPreservesExistingOutput)
		{
			const auto raw = CreateFileW(m_input.c_str(), GENERIC_READ, 0, nullptr, OPEN_EXISTING, 0, nullptr);
			Assert::IsTrue(raw != INVALID_HANDLE_VALUE);
			const Handle lock{raw};
			CheckFailure(Parse(true));
			Assert::AreEqual(0UL, m_reads);
		}

		TEST_METHOD(FailedSizeQueryDoesNotReadParseOrReplaceOutput)
		{
			m_operations.size = [](FILE*) { return __int64{-1}; };
			CheckFailure(Parse(true));
			Assert::AreEqual(0UL, m_reads);
			Assert::IsTrue(Read(m_input) == m_bytes);
		}

		TEST_METHOD(UnsupportedSizeIsRejectedBeforeReadingOrParsing)
		{
			m_operations.size = [](FILE*) { return static_cast<__int64>((std::numeric_limits<ULONG>::max)()) + 1; };
			CheckFailure(Parse(true));
			Assert::AreEqual(0UL, m_reads);
		}

		TEST_METHOD(FailedInputDoesNotCreateOutput)
		{
			Assert::IsTrue(DeleteFileW(m_output.c_str()) != FALSE);
			m_operations.size = [](FILE*) { return __int64{-1}; };
			Assert::IsTrue(FAILED(Parse(true)));
			Assert::AreEqual(INVALID_FILE_ATTRIBUTES, GetFileAttributesW(m_output.c_str()));
			Assert::AreEqual(0UL, m_parses);
		}

		TEST_METHOD(ShortReadDoesNotParseOrReplaceOutput)
		{
			m_operations.read = [](BYTE* bytes, size_t size, FILE* file) { return fread(bytes, 1, size - 1, file); };
			CheckFailure(Parse(true));
			Assert::IsTrue(Read(m_input) == m_bytes);
		}

		TEST_METHOD(ReadErrorDoesNotParseOrReplaceOutput)
		{
			m_operations.error = [](FILE*) { return true; };
			CheckFailure(Parse(false));
		}

		TEST_METHOD(ReadFailureDoesNotCreateOutput)
		{
			Assert::IsTrue(DeleteFileW(m_output.c_str()) != FALSE);
			m_operations.read = [](BYTE*, size_t, FILE*) { return size_t{}; };
			Assert::IsTrue(FAILED(Parse(true)));
			Assert::AreEqual(INVALID_FILE_ATTRIBUTES, GetFileAttributesW(m_output.c_str()));
			Assert::AreEqual(0UL, m_parses);
		}

		TEST_METHOD(AllocationFailurePreservesOutput)
		{
			m_operations.read = [](BYTE*, size_t, FILE*) -> size_t { throw std::bad_alloc{}; };
			CheckFailure(Parse(true));
		}

		TEST_METHOD(SizeBoundariesDoNotAllocateInputBuffers)
		{
			const auto sizeMax = (std::numeric_limits<size_t>::max)();
			const auto ulongMax = static_cast<unsigned __int64>((std::numeric_limits<ULONG>::max)());
			size_t size{42};
			Assert::IsTrue(FAILED(smartviewfile::ValidateInputSize(-1, sizeMax, size)));
			Assert::AreEqual(size_t{}, size);
			Assert::AreEqual(S_OK, smartviewfile::ValidateInputSize(0, 1, size));
			Assert::AreEqual(size_t{}, size);
			Assert::IsTrue(FAILED(smartviewfile::ValidateInputSize(0, 0, size)));
			Assert::AreEqual(S_OK, smartviewfile::ValidateInputSize(7, 8, size));
			Assert::AreEqual(size_t{7}, size);
			Assert::IsTrue(FAILED(smartviewfile::ValidateInputSize(8, 8, size)));
			Assert::AreEqual(size_t{}, size);
			Assert::IsTrue(FAILED(smartviewfile::ValidateInputSize(static_cast<__int64>(ulongMax + 1), sizeMax, size)));
			const auto largest = (std::min)(ulongMax, static_cast<unsigned __int64>(sizeMax - 1));
			Assert::AreEqual(S_OK, smartviewfile::ValidateInputSize(static_cast<__int64>(largest), sizeMax, size));
			Assert::AreEqual(static_cast<size_t>(largest), size);
			if constexpr (sizeof(size_t) == sizeof(ULONG))
				Assert::IsTrue(FAILED(smartviewfile::ValidateInputSize(static_cast<__int64>(sizeMax), sizeMax, size)));
			const auto bufferMax = std::vector<BYTE>{}.max_size();
			const auto largestBuffer = (std::min)(ulongMax, static_cast<unsigned __int64>(bufferMax - 1));
			Assert::AreEqual(
				S_OK, smartviewfile::ValidateInputSize(static_cast<__int64>(largestBuffer), bufferMax, size));
			Assert::IsTrue(
				FAILED(smartviewfile::ValidateInputSize(static_cast<__int64>(largestBuffer + 1), bufferMax, size)));
		}

		TEST_METHOD(SamePathCannotTruncateInput)
		{
			Assert::IsTrue(FAILED(smartviewfile::ParseFile(m_input, m_input, parserType::ENTRYID, true, m_operations)));
			Assert::IsTrue(Read(m_input) == m_bytes);
		}

		TEST_METHOD(RelativePathAliasCannotTruncateInput)
		{
			const auto path = std::filesystem::path(m_input);
			const auto alias = path.parent_path() / L"." / path.filename();
			Assert::IsTrue(
				FAILED(smartviewfile::ParseFile(m_input, alias.wstring(), parserType::ENTRYID, true, m_operations)));
			Assert::IsTrue(Read(m_input) == m_bytes);
		}

		TEST_METHOD(HardLinkAliasCannotTruncateInput)
		{
			m_alias = m_input + L".link";
			Assert::IsTrue(CreateHardLinkW(m_alias.c_str(), m_input.c_str(), nullptr) != FALSE);
			Assert::IsTrue(FAILED(smartviewfile::ParseFile(m_input, m_alias, parserType::ENTRYID, true, m_operations)));
			Assert::IsTrue(Read(m_input) == m_bytes);
			Assert::IsTrue(Read(m_alias) == m_bytes);
		}

		TEST_METHOD(OutputOpenFailurePreservesInput)
		{
			const auto raw = CreateFileW(m_output.c_str(), GENERIC_READ, 0, nullptr, OPEN_EXISTING, 0, nullptr);
			Assert::IsTrue(raw != INVALID_HANDLE_VALUE);
			const Handle lock{raw};
			Assert::IsTrue(FAILED(Parse(true)));
			Assert::IsTrue(Read(m_input) == m_bytes);
		}

		TEST_METHOD(CommandLineReportsInputFailureWithoutReplacingOutput)
		{
			Assert::IsTrue(DeleteFileW(m_input.c_str()) != FALSE);
			Assert::AreEqual(1UL, RunCommand(SmartViewArguments()));
			Assert::IsTrue(Read(m_output) == m_sentinel);
		}

		TEST_METHOD(CommandLineReportsUnreadableInputWithoutReplacingOutput)
		{
			const auto raw = CreateFileW(m_input.c_str(), GENERIC_READ, 0, nullptr, OPEN_EXISTING, 0, nullptr);
			Assert::IsTrue(raw != INVALID_HANDLE_VALUE);
			const Handle lock{raw};
			Assert::AreEqual(1UL, RunCommand(SmartViewArguments()));
			Assert::IsTrue(Read(m_output) == m_sentinel);
		}

		TEST_METHOD(CommandLineValidBinaryAndHexRetainParsedOutput)
		{
			Assert::AreEqual(0UL, RunCommand(SmartViewArguments()));
			CheckOutput(m_bytes);
			const auto text = strings::wstringTostring(strings::BinToHexString(m_bytes, false));
			Write(m_input, std::vector<BYTE>(text.begin(), text.end()));
			Assert::AreEqual(0UL, RunCommand(SmartViewArguments(false)));
			CheckOutput(m_bytes);
		}

		TEST_METHOD(CommandLineRejectsSameFileWithoutChangingInput)
		{
			Assert::AreEqual(
				1UL, RunCommand(L"-ParserType 1 -Binary -Input \"" + m_input + L"\" -Output \"" + m_input + L"\""));
			Assert::IsTrue(Read(m_input) == m_bytes);
		}

		TEST_METHOD(CommandLineEmptyBinaryAndHexInputsSucceed)
		{
			Write(m_input, {});
			for (const auto binary : {false, true})
			{
				Assert::AreEqual(0UL, RunCommand(SmartViewArguments(binary)));
				CheckOutput({});
			}
		}

		TEST_METHOD(CommandLineRejectsUnsupportedSparseFileSizeWithoutReplacingOutput)
		{
			{
				const auto raw =
					CreateFileW(m_input.c_str(), GENERIC_READ | GENERIC_WRITE, 0, nullptr, OPEN_EXISTING, 0, nullptr);
				Assert::IsTrue(raw != INVALID_HANDLE_VALUE);
				const Handle file{raw};
				DWORD returned{};
				Assert::IsTrue(
					DeviceIoControl(file.get(), FSCTL_SET_SPARSE, nullptr, 0, nullptr, 0, &returned, nullptr) != FALSE);
				LARGE_INTEGER length{};
				length.QuadPart = static_cast<__int64>((std::numeric_limits<ULONG>::max)()) + 1;
				Assert::IsTrue(SetFilePointerEx(file.get(), length, nullptr, FILE_BEGIN) != FALSE);
				Assert::IsTrue(SetEndOfFile(file.get()) != FALSE);
			}
			Assert::AreEqual(1UL, RunCommand(SmartViewArguments()));
			Assert::IsTrue(Read(m_output) == m_sentinel);
		}

		TEST_METHOD(UnrelatedCommandRetainsSuccessfulExitCode)
		{
			Assert::AreEqual(0UL, RunCommand(L"-Error 0x8004010F"));
		}

	private:
		static std::wstring TemporaryFile()
		{
			wchar_t temporary[MAX_PATH]{};
			wchar_t file[MAX_PATH]{};
			const auto length = GetTempPathW(_countof(temporary), temporary);
			Assert::IsTrue(length && length < _countof(temporary));
			Assert::IsTrue(GetTempFileNameW(temporary, L"msv", 0, file) != 0);
			return file;
		}

		static void Write(const std::wstring& path, const std::vector<BYTE>& bytes)
		{
			File file{output::MyOpenFileMode(path, L"wb"), output::CloseFile};
			Assert::IsNotNull(file.get());
			Assert::AreEqual(bytes.size(), fwrite(bytes.data(), 1, bytes.size(), file.get()));
		}

		static std::vector<BYTE> Read(const std::wstring& path)
		{
			File file{output::MyOpenFileMode(path, L"rb"), output::CloseFile};
			Assert::IsNotNull(file.get());
			const auto length = _filelengthi64(_fileno(file.get()));
			Assert::IsTrue(length >= 0 && length < 65536);
			std::vector<BYTE> bytes(static_cast<size_t>(length));
			Assert::AreEqual(bytes.size(), fread(bytes.data(), 1, bytes.size(), file.get()));
			return bytes;
		}

		HRESULT Parse(bool binary)
		{
			return smartviewfile::ParseFile(m_input, m_output, parserType::ENTRYID, binary, m_operations);
		}

		void CheckFailure(HRESULT result)
		{
			Assert::IsTrue(FAILED(result));
			Assert::AreEqual(0UL, m_parses);
			Assert::IsTrue(Read(m_output) == m_sentinel);
		}

		void CheckOutput(const std::vector<BYTE>& bytes)
		{
			const auto expected =
				smartview::InterpretBinary(
					{static_cast<ULONG>(bytes.size()), const_cast<BYTE*>(bytes.data())}, parserType::ENTRYID, nullptr)
					->toString();
			const auto actual = Read(m_output);
			const auto begin = reinterpret_cast<const BYTE*>(expected.data());
			const std::vector<BYTE> expectedBytes(begin, begin + expected.size() * sizeof(wchar_t));
			Assert::IsTrue(
				actual == expectedBytes, L"Parsed output must retain its existing UTF-16 byte representation.");
		}

		std::wstring SmartViewArguments(bool binary = true) const
		{
			size_t parser{};
			for (size_t index = 1; index < SmartViewParserTypeArray.size(); ++index)
				if (SmartViewParserTypeArray[index].type == parserType::ENTRYID) parser = index;
			Assert::IsTrue(parser != 0);
			return L"-ParserType " + std::to_wstring(parser) + (binary ? L" -Binary" : L"") + L" -Input \"" + m_input +
				   L"\" -Output \"" + m_output + L"\"";
		}

		static DWORD RunCommand(const std::wstring& arguments)
		{
			wchar_t module[MAX_PATH]{};
			Assert::IsTrue(GetModuleFileNameW(GetModuleHandleW(L"UnitTest.dll"), module, _countof(module)) != 0);
			const auto directory = std::filesystem::path(module).parent_path();
			const auto executable =
				directory.parent_path().parent_path() / L"MrMAPI" / directory.filename() / L"MrMAPI.exe";
			auto command = L"\"" + executable.wstring() + L"\" -NoAddins " + arguments;
			SECURITY_ATTRIBUTES attributes{sizeof(attributes), nullptr, TRUE};
			const auto raw = CreateFileW(
				L"NUL",
				GENERIC_READ | GENERIC_WRITE,
				FILE_SHARE_READ | FILE_SHARE_WRITE,
				&attributes,
				OPEN_EXISTING,
				0,
				nullptr);
			Assert::IsTrue(raw != INVALID_HANDLE_VALUE);
			const Handle console{raw};
			STARTUPINFOW startup{};
			startup.cb = sizeof(startup);
			startup.dwFlags = STARTF_USESTDHANDLES;
			startup.hStdInput = console.get();
			startup.hStdOutput = console.get();
			startup.hStdError = console.get();
			PROCESS_INFORMATION process{};
			Assert::IsTrue(
				CreateProcessW(
					nullptr,
					command.data(),
					nullptr,
					nullptr,
					TRUE,
					CREATE_NO_WINDOW,
					nullptr,
					nullptr,
					&startup,
					&process) != FALSE);
			const Handle thread{process.hThread};
			const Handle child{process.hProcess};
			Assert::AreEqual(static_cast<DWORD>(WAIT_OBJECT_0), WaitForSingleObject(child.get(), 30000));
			DWORD exitCode{};
			Assert::IsTrue(GetExitCodeProcess(child.get(), &exitCode) != FALSE);
			return exitCode;
		}

		std::wstring m_input;
		std::wstring m_output;
		std::wstring m_alias;
		std::vector<BYTE> m_bytes;
		const std::vector<BYTE> m_sentinel{'p', 'r', 'e', 's', 'e', 'r', 'v', 'e'};
		smartviewfile::InputOperations m_operations;
		ULONG m_reads{};
		ULONG m_parses{};
		std::vector<BYTE> m_parsed;
	};
} // namespace mmsmartviewtest
