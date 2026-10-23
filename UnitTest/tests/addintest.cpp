#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/addin/addin.h>
#include <core/addin/addinfile.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/registry.h>
#include <cstring>
#include <stdexcept>

namespace addin
{
	// imports for testing
	int _cdecl compareTypes(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept;
	int _cdecl compareTags(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept;
	int _cdecl compareNameID(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept;
	int _cdecl compareSmartViewParser(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept;

	std::wstring AddInStructTypeToString(parserType parser);

	void SortFlagArray(_In_count_(ulFlags) LPFLAG_ARRAY_ENTRY lpFlags, _In_ size_t ulFlags) noexcept;
	void AppendFlagIfNotDupe(std::vector<FLAG_ARRAY_ENTRY>& target, FLAG_ARRAY_ENTRY source);

	void MergeFlagArrays(std::vector<FLAG_ARRAY_ENTRY>& In1, _In_count_(cIn2) LPFLAG_ARRAY_ENTRY In2, _In_ size_t cIn2);
} // namespace addin

namespace addintest
{
	class TestDirectory
	{
	public:
		TestDirectory()
		{
			WCHAR temporary[MAX_PATH]{};
			WCHAR name[MAX_PATH]{};
			if (!GetTempPathW(_countof(temporary), temporary) || !GetTempFileNameW(temporary, L"adi", 0, name))
				throw std::runtime_error("Unable to generate add-in test directory");
			path = name;
			if (!DeleteFileW(path.c_str()) || !CreateDirectoryW(path.c_str(), nullptr))
				throw std::runtime_error("Unable to create add-in test directory");
		}

		~TestDirectory()
		{
			for (const auto& file : files)
				if (!DeleteFileW(file.c_str())) Logger::WriteMessage(L"Unable to remove add-in test file");
			if (!RemoveDirectoryW((path + L"\\nested").c_str()) && GetLastError() != ERROR_FILE_NOT_FOUND)
				Logger::WriteMessage(L"Unable to remove nested add-in test directory");
			if (!RemoveDirectoryW(path.c_str())) Logger::WriteMessage(L"Unable to remove add-in test directory");
		}

		void Write(const std::wstring& name, const std::vector<BYTE>& image)
		{
			const auto file = path + L"\\" + name;
			const auto handle =
				CreateFileW(file.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
			if (handle == INVALID_HANDLE_VALUE) throw std::runtime_error("Unable to create add-in test file");
			files.push_back(file);
			DWORD written{};
			const auto success = WriteFile(handle, image.data(), static_cast<DWORD>(image.size()), &written, nullptr);
			const auto closed = CloseHandle(handle);
			Assert::IsTrue(success && closed);
			Assert::AreEqual(static_cast<DWORD>(image.size()), written);
		}

		std::wstring path;

	private:
		std::vector<std::wstring> files;
	};

	class TestRegistry
	{
	public:
		TestRegistry() : savedKeys(registry::RegKeys)
		{
			static LONG sequence{};
			name = L"Software\\MFCMAPI-AddinTest-" + std::to_wstring(GetCurrentProcessId()) + L"-" +
				   std::to_wstring(GetTickCount64()) + L"-" + std::to_wstring(InterlockedIncrement(&sequence));
			DWORD disposition{};
			if (RegCreateKeyExW(
					HKEY_CURRENT_USER,
					name.c_str(),
					0,
					nullptr,
					REG_OPTION_NON_VOLATILE,
					KEY_ALL_ACCESS,
					nullptr,
					&key,
					&disposition) != ERROR_SUCCESS)
				throw std::runtime_error("Unable to create add-in test registry key");
			if (disposition != REG_CREATED_NEW_KEY)
			{
				RegCloseKey(key);
				throw std::runtime_error("Add-in test registry key already exists");
			}
			if (RegOverridePredefKey(HKEY_CURRENT_USER, key) != ERROR_SUCCESS)
			{
				RegCloseKey(key);
				RegDeleteKeyW(HKEY_CURRENT_USER, name.c_str());
				throw std::runtime_error("Unable to isolate add-in test registry");
			}
			savedDirectory = registry::addInDirectory.szCurSTRING;
			registry::addInDirectory.clear();
		}

		~TestRegistry()
		{
			registry::RegKeys = std::move(savedKeys);
			registry::addInDirectory = savedDirectory;
			if (RegOverridePredefKey(HKEY_CURRENT_USER, nullptr) != ERROR_SUCCESS)
				Logger::WriteMessage(L"Unable to restore predefined registry key");
			if (RegDeleteTreeW(key, nullptr) != ERROR_SUCCESS)
				Logger::WriteMessage(L"Unable to clear add-in test registry");
			RegCloseKey(key);
			if (RegDeleteKeyW(HKEY_CURRENT_USER, name.c_str()) != ERROR_SUCCESS)
				Logger::WriteMessage(L"Unable to remove add-in test registry key");
		}

		void SetRaw(DWORD type, const BYTE* data, DWORD size)
		{
			Assert::AreEqual(
				static_cast<LSTATUS>(ERROR_SUCCESS), RegSetValueExW(key, L"AddInDirectory", 0, type, data, size));
		}

		HKEY key{};

	private:
		std::wstring name;
		std::wstring savedDirectory;
		std::vector<registry::__RegKey*> savedKeys;
	};

	class TestWorkingDirectory
	{
	public:
		explicit TestWorkingDirectory(const std::wstring& directory)
		{
			WCHAR current[MAX_PATH]{};
			const auto length = GetCurrentDirectoryW(_countof(current), current);
			if (!length || length >= _countof(current))
				throw std::runtime_error("Unable to save test working directory");
			saved = current;
			if (!SetCurrentDirectoryW(directory.c_str()))
				throw std::runtime_error("Unable to set test working directory");
		}

		~TestWorkingDirectory()
		{
			if (!SetCurrentDirectoryW(saved.c_str())) Logger::WriteMessage(L"Unable to restore test working directory");
		}

	private:
		std::wstring saved;
	};

	void CopyToImage(std::vector<BYTE>& image, const size_t offset, const void* source, const size_t size)
	{
		if (offset > image.size() || size > image.size() - offset || (size && !source))
		{
			throw std::out_of_range("Invalid synthetic PE image copy");
		}

		const auto result = memcpy_s(image.data() + offset, image.size() - offset, source, size);
		if (result != 0) throw std::runtime_error("Unable to copy synthetic PE image data");
	}

	int _cdecl testCompareTypes(_In_ const NAME_ARRAY_ENTRY& a1, _In_ const NAME_ARRAY_ENTRY& a2) noexcept
	{
		return addin::compareTypes(nullptr, reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	int _cdecl testCompareTags(_In_ const NAME_ARRAY_ENTRY_V2& a1, _In_ const NAME_ARRAY_ENTRY_V2& a2) noexcept
	{
		return addin::compareTags(nullptr, reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	int _cdecl testCompareNameID(_In_ const NAMEID_ARRAY_ENTRY& a1, _In_ const NAMEID_ARRAY_ENTRY& a2) noexcept
	{
		return addin::compareNameID(nullptr, reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	int _cdecl testCompareSmartViewParser(
		_In_ const SMARTVIEW_PARSER_ARRAY_ENTRY& a1,
		_In_ const SMARTVIEW_PARSER_ARRAY_ENTRY& a2) noexcept
	{
		return addin::compareSmartViewParser(nullptr, reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	std::vector<BYTE> MakePeImageWithExport(const char* exportName, bool is64Bit = true)
	{
		constexpr size_t ntOffset = 0x80;
		constexpr size_t rawOffset = 0x200;
		auto image = std::vector<BYTE>(0x400);

		IMAGE_DOS_HEADER dosHeader = {};
		dosHeader.e_magic = IMAGE_DOS_SIGNATURE;
		dosHeader.e_lfanew = static_cast<LONG>(ntOffset);
		CopyToImage(image, 0, &dosHeader, sizeof(dosHeader));

		const DWORD signature = IMAGE_NT_SIGNATURE;
		CopyToImage(image, ntOffset, &signature, sizeof(signature));

		IMAGE_FILE_HEADER fileHeader = {};
		fileHeader.Machine = is64Bit ? IMAGE_FILE_MACHINE_AMD64 : IMAGE_FILE_MACHINE_I386;
		fileHeader.NumberOfSections = 1;
		fileHeader.SizeOfOptionalHeader =
			static_cast<WORD>(is64Bit ? sizeof(IMAGE_OPTIONAL_HEADER64) : sizeof(IMAGE_OPTIONAL_HEADER32));
		CopyToImage(image, ntOffset + sizeof(signature), &fileHeader, sizeof(fileHeader));

		const auto optionalOffset = ntOffset + sizeof(signature) + sizeof(fileHeader);
		if (is64Bit)
		{
			IMAGE_OPTIONAL_HEADER64 optionalHeader = {};
			optionalHeader.Magic = IMAGE_NT_OPTIONAL_HDR64_MAGIC;
			optionalHeader.SizeOfHeaders = static_cast<DWORD>(rawOffset);
			optionalHeader.NumberOfRvaAndSizes = IMAGE_NUMBEROF_DIRECTORY_ENTRIES;
			optionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT] =
				IMAGE_DATA_DIRECTORY{0x1000, sizeof(IMAGE_EXPORT_DIRECTORY)};
			CopyToImage(image, optionalOffset, &optionalHeader, sizeof(optionalHeader));
		}
		else
		{
			IMAGE_OPTIONAL_HEADER32 optionalHeader = {};
			optionalHeader.Magic = IMAGE_NT_OPTIONAL_HDR32_MAGIC;
			optionalHeader.SizeOfHeaders = static_cast<DWORD>(rawOffset);
			optionalHeader.NumberOfRvaAndSizes = IMAGE_NUMBEROF_DIRECTORY_ENTRIES;
			optionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT] =
				IMAGE_DATA_DIRECTORY{0x1000, sizeof(IMAGE_EXPORT_DIRECTORY)};
			CopyToImage(image, optionalOffset, &optionalHeader, sizeof(optionalHeader));
		}

		IMAGE_SECTION_HEADER section = {};
		section.VirtualAddress = 0x1000;
		section.SizeOfRawData = 0x200;
		section.PointerToRawData = static_cast<DWORD>(rawOffset);
		const auto sectionOffset = optionalOffset + fileHeader.SizeOfOptionalHeader;
		CopyToImage(image, sectionOffset, &section, sizeof(section));

		IMAGE_EXPORT_DIRECTORY exports = {};
		exports.NumberOfNames = 1;
		exports.AddressOfNames = 0x1040;
		CopyToImage(image, rawOffset, &exports, sizeof(exports));

		const DWORD nameRva = 0x1060;
		CopyToImage(image, rawOffset + 0x40, &nameRva, sizeof(nameRva));

		const auto exportNameOffset = rawOffset + 0x60;
		const auto exportNameCapacity = image.size() - exportNameOffset;
		if (!exportName) throw std::invalid_argument("Synthetic PE export name cannot be null");
		const auto exportNameLength = strnlen_s(exportName, exportNameCapacity);
		if (exportNameLength == exportNameCapacity)
		{
			throw std::length_error("Synthetic PE export name exceeds the image");
		}

		CopyToImage(image, exportNameOffset, exportName, exportNameLength + 1);
		return image;
	}

	TEST_CLASS(addinTest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_AddInDirectoryValidation)
		{
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::ValidateAddInDirectory(L""));
			for (const auto& invalid :
				 {L"addins",
				  L".\\addins",
				  L"C:addins",
				  L"\\addins",
				  L"\\\\server\\share",
				  L"\\\\?\\C:\\addins",
				  L"\\\\.\\C:\\addins",
				  L"\"C:\\addins\"",
				  L"C:\\*.dll",
				  L"C:/addins",
				  L"%TEMP%\\addins",
				  L"C:\\addins\\..",
				  L"C:\\addins\\.",
				  L"C:\\addins ",
				  L"C:\\addins|other",
				  L"C:\\addins:other"})
				Assert::AreNotEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::ValidateAddInDirectory(invalid), invalid);
			const auto embeddedNull = std::wstring(L"C:\\addins") + L'\0' + L"ignored";
			Assert::AreNotEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::ValidateAddInDirectory(embeddedNull));
			Assert::AreEqual(
				static_cast<DWORD>(ERROR_FILENAME_EXCED_RANGE),
				addin::ValidateAddInDirectory(L"C:\\" + std::wstring(MAX_PATH, L'a')));

			TestDirectory directory;
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::ValidateAddInDirectory(directory.path));
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::ValidateAddInDirectory(directory.path + L"\\"));
			Assert::IsTrue(CreateDirectoryW((directory.path + L"\\nested").c_str(), nullptr));
			Assert::AreEqual(
				static_cast<DWORD>(ERROR_SUCCESS), addin::ValidateAddInDirectory(directory.path + L"\\nested"));
			const auto unicodePath = directory.path + L"\\spaces ; \u65e5\u672c";
			Assert::IsTrue(CreateDirectoryW(unicodePath.c_str(), nullptr));
			const auto unicodeResult = addin::ValidateAddInDirectory(unicodePath);
			Assert::IsTrue(RemoveDirectoryW(unicodePath.c_str()));
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), unicodeResult);
			directory.Write(L"not a directory.dll", {0});
			Assert::AreEqual(
				static_cast<DWORD>(ERROR_DIRECTORY),
				addin::ValidateAddInDirectory(directory.path + L"\\not a directory.dll"));
			Assert::AreNotEqual(
				static_cast<DWORD>(ERROR_SUCCESS), addin::ValidateAddInDirectory(directory.path + L"\\missing"));
		}

		TEST_METHOD(Test_AddInDirectoryRegistry)
		{
			TestRegistry registry;
			registry::ReadFromRegistry();
			Assert::IsTrue(registry::addInDirectory.empty());

			const std::wstring directory = L"C:\\Add-ins with spaces;\u65e5\u672c";
			registry::addInDirectory = directory;
			registry::WriteToRegistry();
			registry::addInDirectory.clear();
			registry::ReadFromRegistry();
			Assert::AreEqual(directory, registry::addInDirectory.szCurSTRING);
			registry::addInDirectory.clear();
			registry::WriteToRegistry();
			const auto root = registry::CreateRootKey();
			Assert::IsNotNull(root);
			const auto queryResult = RegQueryValueExW(root, L"AddInDirectory", nullptr, nullptr, nullptr, nullptr);
			RegCloseKey(root);
			Assert::AreEqual(static_cast<LSTATUS>(ERROR_FILE_NOT_FOUND), queryResult);
			registry::ReadFromRegistry();
			Assert::IsTrue(registry::addInDirectory.empty());
		}

		TEST_METHOD(Test_AddInDirectoryMalformedRegistry)
		{
			TestRegistry registry;
			const DWORD wrongType = 1;
			registry.SetRaw(REG_DWORD, reinterpret_cast<const BYTE*>(&wrongType), sizeof(wrongType));
			Assert::IsTrue(registry::ReadStringFromRegistry(registry.key, L"AddInDirectory").empty());
			const std::wstring directory = L"C:\\addins";
			const std::wstring embeddedNull = directory + L'\0' + L"ignored";
			registry.SetRaw(
				REG_SZ,
				reinterpret_cast<const BYTE*>(embeddedNull.c_str()),
				static_cast<DWORD>((embeddedNull.size() + 1) * sizeof(wchar_t)));
			const auto value = registry::ReadStringFromRegistry(registry.key, L"AddInDirectory");
			Assert::AreEqual(embeddedNull, value);
			Assert::AreNotEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::ValidateAddInDirectory(value));
			registry.SetRaw(REG_SZ, reinterpret_cast<const BYTE*>(directory.data()), 3);
			Assert::IsTrue(registry::ReadStringFromRegistry(registry.key, L"AddInDirectory").empty());
			registry.SetRaw(
				REG_EXPAND_SZ,
				reinterpret_cast<const BYTE*>(directory.c_str()),
				static_cast<DWORD>((directory.size() + 1) * sizeof(wchar_t)));
			Assert::IsTrue(registry::ReadStringFromRegistry(registry.key, L"AddInDirectory").empty());
			const std::wstring longDirectory = L"C:\\" + std::wstring(MAX_PATH, L'a');
			registry::WriteStringToRegistry(registry.key, L"AddInDirectory", longDirectory);
			const auto longValue = registry::ReadStringFromRegistry(registry.key, L"AddInDirectory");
			Assert::AreEqual(longDirectory, longValue);
			Assert::AreEqual(static_cast<DWORD>(ERROR_FILENAME_EXCED_RANGE), addin::ValidateAddInDirectory(longValue));
		}

		TEST_METHOD(Test_AddInsDisabledByDefault)
		{
			TestRegistry registry;
			const auto root = registry::CreateRootKey();
			Assert::IsNotNull(root);
			registry::WriteDWORDToRegistry(root, L"LoadAddIns", 1);
			RegCloseKey(root);
			registry::ReadFromRegistry();
			bool prompted{};
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns([&prompted](const auto&, auto&) {
								 prompted = true;
								 return false;
							 }));
			Assert::IsFalse(prompted);
			Assert::IsTrue(g_lpMyAddins.empty());
			Assert::IsFalse(PropTagArray.empty());
			const auto report = addin::GetAddInReport();
			Assert::IsTrue(report.find(strings::loadstring(IDS_ADDINREPORTDISABLED)) != std::wstring::npos);
			Assert::IsTrue(report.find(strings::loadstring(IDS_ADDINREPORTNOFILES)) != std::wstring::npos);
		}

		TEST_METHOD(Test_AddInsOnlyScanConfiguredDirectory)
		{
			TestRegistry registry;
			TestDirectory directory;
			TestDirectory otherDirectory;
			TestWorkingDirectory workingDirectory(otherDirectory.path);
			const auto root = registry::CreateRootKey();
			Assert::IsNotNull(root);
			registry::WriteDWORDToRegistry(root, L"LoadAddIns", 0);
			RegCloseKey(root);
			Assert::IsTrue(CreateDirectoryW((directory.path + L"\\nested").c_str(), nullptr));
			auto image = MakePeImageWithExport("LoadAddIn", sizeof(void*) == 8);
			directory.Write(L"selected.dll", image);
			// Distinct hashes keep prompt deduplication from hiding an extra scan.
			image.back() = 1;
			otherDirectory.Write(L"working-directory-excluded.dll", image);
			image.back() = 2;
			directory.Write(L"nested\\excluded.dll", image);
			directory.Write(L"not-an-addin.dll", MakePeImageWithExport("OtherExport", sizeof(void*) == 8));
			registry::addInDirectory = directory.path + L"\\";
			size_t count{};
			Assert::AreEqual(
				static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns([&count](const auto& candidates, auto&) {
					count = candidates.size();
					Assert::AreEqual(std::wstring(L"selected.dll"), candidates[0].fileName);
					return false;
				}));
			Assert::AreEqual(size_t{1}, count);
			Assert::IsTrue(g_lpMyAddins.empty());
			Assert::IsFalse(PropTagArray.empty());
			const auto report = addin::GetAddInReport();
			Assert::IsTrue(report.find(L"selected.dll") != std::wstring::npos);
			Assert::IsTrue(report.find(L"not-an-addin.dll") != std::wstring::npos);
			Assert::IsTrue(report.find(strings::loadstring(IDS_ADDINREPORTPENDING)) != std::wstring::npos);
			Assert::IsTrue(report.find(strings::loadstring(IDS_ADDINREPORTNOTADDIN)) != std::wstring::npos);
			Assert::IsTrue(report.find(L"excluded.dll") == std::wstring::npos);
		}

		TEST_METHOD(Test_AddInsUnknownAndDenied)
		{
			TestRegistry registry;
			TestDirectory directory;
			TestDirectory movedDirectory;
			const auto image = MakePeImageWithExport("LoadAddIn", sizeof(void*) == 8);
			directory.Write(L"unknown.dll", image);
			movedDirectory.Write(L"moved.dll", image);
			registry::addInDirectory = directory.path;
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns());
			Assert::IsTrue(g_lpMyAddins.empty());
			Assert::IsTrue(
				addin::GetAddInReport().find(strings::loadstring(IDS_ADDINREPORTPENDING)) != std::wstring::npos);

			size_t prompts{};
			const addin::AddInApprovalPrompt deny = [&prompts](const auto& candidates, auto& decisions) {
				++prompts;
				decisions.assign(candidates.size(), false);
				return true;
			};
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns(deny));
			Assert::AreEqual(size_t{1}, prompts);
			registry::addInDirectory = movedDirectory.path;
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns(deny));
			Assert::AreEqual(size_t{1}, prompts);
			Assert::IsTrue(g_lpMyAddins.empty());
			const auto report = addin::GetAddInReport();
			Assert::IsTrue(report.find(L"moved.dll") != std::wstring::npos);
			Assert::IsTrue(report.find(L"unknown.dll") == std::wstring::npos);
			Assert::IsTrue(report.find(strings::loadstring(IDS_ADDINREPORTDENIED)) != std::wstring::npos);
		}

		TEST_METHOD(Test_AddInReportShowsSkippedBitnessAndLoadFailure)
		{
			TestRegistry registry;
			TestDirectory directory;
			directory.Write(L"wrong-bitness.dll", MakePeImageWithExport("LoadAddIn", sizeof(void*) != 8));
			directory.Write(L"invalid-load.dll", MakePeImageWithExport("LoadAddIn", sizeof(void*) == 8));
			registry::addInDirectory = directory.path;
			size_t prompted{};
			Assert::AreEqual(
				static_cast<DWORD>(ERROR_SUCCESS),
				addin::LoadAddIns([&prompted](const auto& candidates, auto& decisions) {
					prompted = candidates.size();
					decisions.assign(candidates.size(), true);
					return true;
				}));
			Assert::AreEqual(size_t{1}, prompted);
			Assert::IsTrue(g_lpMyAddins.empty());
			const auto report = addin::GetAddInReport();
			Assert::IsTrue(report.find(L"wrong-bitness.dll") != std::wstring::npos);
			Assert::IsTrue(report.find(strings::loadstring(IDS_ADDINREPORTBITNESS)) != std::wstring::npos);
			Assert::IsTrue(report.find(L"invalid-load.dll") != std::wstring::npos);
			Assert::IsTrue(report.find(strings::loadstring(IDS_ADDINREPORTLOADFAILED)) != std::wstring::npos);
			Assert::IsTrue(report.find(L"Windows error:") != std::wstring::npos);
			Assert::IsTrue(report.find(L"\r\n  Loaded.") == std::wstring::npos);
		}

		TEST_METHOD(Test_AddInReportDoesNotScanChangedConfiguration)
		{
			TestRegistry registry;
			TestDirectory directory;
			registry::addInDirectory = directory.path;
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns());
			directory.Write(L"not-scanned.dll", MakePeImageWithExport("LoadAddIn", sizeof(void*) == 8));
			registry::addInDirectory.clear();
			const auto report = addin::GetAddInReport();
			Assert::IsTrue(
				report.find(strings::formatmessage(
					IDS_ADDINREPORTCONFIGURED, strings::loadstring(IDS_ADDINREPORTDISABLED).c_str())) !=
				std::wstring::npos);
			Assert::IsTrue(
				report.find(strings::formatmessage(IDS_ADDINREPORTSCANNED, directory.path.c_str())) !=
				std::wstring::npos);
			Assert::IsTrue(report.find(L"not-scanned.dll") == std::wstring::npos);
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns());
			Assert::IsTrue(addin::GetAddInReport().find(directory.path) == std::wstring::npos);
		}

		TEST_METHOD(Test_AddInReportPreviewsPendingDirectoryWithoutChangingConfiguration)
		{
			TestRegistry registry;
			TestDirectory directory;
			TestDirectory pendingDirectory;
			registry::addInDirectory = directory.path;
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns());
			pendingDirectory.Write(L"not-scanned.dll", MakePeImageWithExport("LoadAddIn", sizeof(void*) == 8));
			const auto originalReport = addin::GetAddInReport();
			const auto report = addin::GetAddInReport(pendingDirectory.path);
			Assert::IsTrue(
				report.find(strings::formatmessage(IDS_ADDINREPORTCONFIGURED, pendingDirectory.path.c_str())) !=
				std::wstring::npos);
			Assert::IsTrue(
				report.find(strings::formatmessage(IDS_ADDINREPORTSCANNED, directory.path.c_str())) !=
				std::wstring::npos);
			Assert::IsTrue(report.find(L"not-scanned.dll") == std::wstring::npos);
			const auto disabledReport = addin::GetAddInReport(L"");
			Assert::IsTrue(
				disabledReport.find(strings::formatmessage(
					IDS_ADDINREPORTCONFIGURED, strings::loadstring(IDS_ADDINREPORTDISABLED).c_str())) !=
				std::wstring::npos);
			Assert::IsTrue(
				disabledReport.find(strings::formatmessage(IDS_ADDINREPORTSCANNED, directory.path.c_str())) !=
				std::wstring::npos);
			Assert::AreEqual(directory.path, registry::addInDirectory.szCurSTRING);
			Assert::AreEqual(originalReport, addin::GetAddInReport());
			Assert::IsTrue(g_lpMyAddins.empty());
		}

		TEST_METHOD(Test_ClearAddInDecisions)
		{
			TestRegistry registry;
			const auto root = registry::CreateRootKey();
			Assert::IsNotNull(root);
			registry::WriteStringToRegistry(root, L"AddInApprovalList", L"approved-hash");
			registry::WriteStringToRegistry(root, L"AddInDenialList", L"denied-hash");
			registry::WriteStringToRegistry(root, L"AddInDirectory", L"C:\\addins");

			Assert::IsTrue(addin::ClearAddInApprovalCache());
			for (const auto name : {L"AddInApprovalList", L"AddInDenialList"})
			{
				Assert::AreEqual(
					static_cast<LSTATUS>(ERROR_FILE_NOT_FOUND),
					RegQueryValueExW(root, name, nullptr, nullptr, nullptr, nullptr));
			}
			Assert::AreEqual(std::wstring(L"C:\\addins"), registry::ReadStringFromRegistry(root, L"AddInDirectory"));
			Assert::IsTrue(addin::ClearAddInApprovalCache());
			RegCloseKey(root);
		}

		TEST_METHOD(Test_OnlyRegisteredOptionsAreRead)
		{
			TestRegistry registry;
			const auto root = registry::CreateRootKey();
			Assert::IsNotNull(root);
			const auto oldDebugTag = registry::debugTag.ulCurDWORD;
			const auto oldGetPropList = registry::useGetPropList.ulCurDWORD;
			registry::WriteDWORDToRegistry(root, L"DebugTag", ~oldDebugTag);
			registry::WriteDWORDToRegistry(root, L"UseGetPropList", !oldGetPropList);
			registry::WriteStringToRegistry(root, L"AddInDirectory", L"C:\\addins");
			RegCloseKey(root);

			registry::RegKeys = {&registry::addInDirectory};
			registry::ReadFromRegistry();
			Assert::AreEqual(std::wstring(L"C:\\addins"), registry::addInDirectory.szCurSTRING);
			Assert::AreEqual(oldDebugTag, registry::debugTag.ulCurDWORD);
			Assert::AreEqual(oldGetPropList, registry::useGetPropList.ulCurDWORD);
		}

		TEST_METHOD(Test_AddInsEmptyAndInvalidDirectories)
		{
			TestRegistry registry;
			TestDirectory directory;
			registry::addInDirectory = directory.path;
			Assert::AreEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns());
			Assert::IsTrue(
				addin::GetAddInReport().find(strings::loadstring(IDS_ADDINREPORTNOFILES)) != std::wstring::npos);
			registry::addInDirectory = directory.path + L"\\missing";
			bool prompted{};
			Assert::AreNotEqual(static_cast<DWORD>(ERROR_SUCCESS), addin::LoadAddIns([&prompted](const auto&, auto&) {
									prompted = true;
									return false;
								}));
			Assert::IsFalse(prompted);
			Assert::IsTrue(g_lpMyAddins.empty());
			Assert::IsFalse(PropTagArray.empty());
			const auto report = addin::GetAddInReport();
			Assert::IsTrue(report.find(directory.path + L"\\missing") != std::wstring::npos);
			Assert::IsTrue(report.find(L"Windows error:") != std::wstring::npos);
		}

		TEST_METHOD(Test_compareTypes)
		{
			Assert::AreEqual(0, testCompareTypes({1, L"one"}, {1, L"one"}));
			Assert::AreEqual(-1, testCompareTypes({1, L"one"}, {2, L"two"}));
			Assert::AreEqual(1, testCompareTypes({2, L"two"}, {1, L"one"}));

			Assert::AreEqual(-1, testCompareTypes({1, L"a"}, {1, L"b"}));
			Assert::AreEqual(1, testCompareTypes({1, L"b"}, {1, L"a"}));

			// Same name gets no special treatment
			Assert::AreEqual(-1, testCompareTypes({1, L"one"}, {2, L"one"}));
			Assert::AreEqual(1, testCompareTypes({2, L"one"}, {1, L"one"}));
		}

		TEST_METHOD(Test_compareTags)
		{
			Assert::AreEqual(0, testCompareTags({1, 0, L"one"}, {1, 0, L"one"}));
			Assert::AreEqual(-1, testCompareTags({1, 0, L"one"}, {2, 0, L"two"}));
			Assert::AreEqual(1, testCompareTags({2, 0, L"two"}, {1, 0, L"one"}));

			Assert::AreEqual(-1, testCompareTags({1, 0, L"a"}, {1, 0, L"b"}));
			Assert::AreEqual(-1, testCompareTags({1, 0, L"a"}, {1, 1, L"b"}));

			// Sort order field doesn't actually affect this particular sort
			Assert::AreEqual(-1, testCompareTags({1, 1, L"a"}, {1, 0, L"b"}));
			Assert::AreEqual(1, testCompareTags({1, 0, L"b"}, {1, 0, L"a"}));

			// Same name compares equal
			Assert::AreEqual(0, testCompareTags({1, 0, L"one"}, {2, 0, L"one"}));
		}

		TEST_METHOD(Test_compareNameID)
		{
			Assert::AreEqual(
				0,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"}));
			Assert::AreEqual(
				-1,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{2, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"}));
			Assert::AreEqual(
				1,
				testCompareNameID(
					{2, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"}));
			Assert::AreEqual(
				-1,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{1, &guid::PSETID_Note, L"a", PT_SYSTIME, L"aa"}));
			Assert::AreEqual(
				-1,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{1, &guid::PSETID_Meeting, L"b", PT_SYSTIME, L"bb"}));
			Assert::AreEqual(
				1,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"b", PT_SYSTIME, L"bb"},
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"}));
		}

		TEST_METHOD(Test_CompareSmartViewParser)
		{
			Assert::AreEqual(
				0, testCompareSmartViewParser({1, parserType::REPORTTAG, false}, {1, parserType::REPORTTAG, false}));
			Assert::AreEqual(
				-1, testCompareSmartViewParser({1, parserType::REPORTTAG, false}, {2, parserType::REPORTTAG, false}));
			Assert::AreEqual(
				1, testCompareSmartViewParser({2, parserType::REPORTTAG, false}, {1, parserType::REPORTTAG, false}));

			Assert::AreEqual(
				-1, testCompareSmartViewParser({1, parserType::ENTRYID, false}, {1, parserType::REPORTTAG, false}));
			Assert::AreEqual(
				1, testCompareSmartViewParser({1, parserType::REPORTTAG, false}, {1, parserType::ENTRYID, false}));

			Assert::AreEqual(
				-1, testCompareSmartViewParser({1, parserType::REPORTTAG, false}, {1, parserType::REPORTTAG, true}));
			Assert::AreEqual(
				1, testCompareSmartViewParser({1, parserType::REPORTTAG, true}, {1, parserType::REPORTTAG, false}));
		}

		//AddInStructTypeToString
		TEST_METHOD(Test_AddInStructTypeToString)
		{
			Assert::AreEqual(
				L"Choose Smart View Parser",
				addin::AddInStructTypeToString(parserType::NOPARSING).c_str(),
				L"NOPARSING");
			Assert::AreEqual(L"", addin::AddInStructTypeToString(parserType{0xfff}).c_str(), L"unknown");
		}

		TEST_METHOD(Test_HasLoadAddInExport)
		{
			auto addInImage = MakePeImageWithExport("LoadAddIn");
			Assert::IsTrue(addin::HasLoadAddInExport(addInImage.data(), addInImage.size()));
			Assert::IsTrue(addin::IsCompatibleBitness(addInImage.data(), addInImage.size(), true));
			Assert::IsFalse(addin::IsCompatibleBitness(addInImage.data(), addInImage.size(), false));

			auto x86Image = MakePeImageWithExport("LoadAddIn", false);
			Assert::IsTrue(addin::HasLoadAddInExport(x86Image.data(), x86Image.size()));
			Assert::IsTrue(addin::IsCompatibleBitness(x86Image.data(), x86Image.size(), false));
			Assert::IsFalse(addin::IsCompatibleBitness(x86Image.data(), x86Image.size(), true));

			auto otherImage = MakePeImageWithExport("NotAnAddIn");
			Assert::IsFalse(addin::HasLoadAddInExport(otherImage.data(), otherImage.size()));

			auto invalidNameRva = MakePeImageWithExport("LoadAddIn");
			const DWORD invalidRva = 0xfffffff0;
			CopyToImage(invalidNameRva, 0x240, &invalidRva, sizeof(invalidRva));
			Assert::IsFalse(addin::HasLoadAddInExport(invalidNameRva.data(), invalidNameRva.size()));

			auto unterminatedName = MakePeImageWithExport("LoadAddIn");
			memset(unterminatedName.data() + 0x260, 'A', unterminatedName.size() - 0x260);
			Assert::IsFalse(addin::HasLoadAddInExport(unterminatedName.data(), unterminatedName.size()));

			addInImage.resize(32);
			Assert::IsFalse(addin::HasLoadAddInExport(addInImage.data(), addInImage.size()));
		}

		void testFA(std::wstring testName, const FLAG_ARRAY_ENTRY& fa1, const FLAG_ARRAY_ENTRY& fa2)
		{
			Assert::AreEqual(fa1.ulFlagName, fa2.ulFlagName, (testName + L"-ulFlagName").c_str());
			Assert::AreEqual(fa1.lFlagValue, fa2.lFlagValue, (testName + L"-lFlagValue").c_str());
			Assert::AreEqual(fa1.ulFlagType, fa2.ulFlagType, (testName + L"-ulFlagType").c_str());
			Assert::AreEqual(fa1.lpszName, fa2.lpszName, (testName + L"-lpszName").c_str());
		}

		TEST_METHOD(Test_FlagArray)
		{
			FLAG_ARRAY_ENTRY flagArray[] = {
				{2, 1, flagVALUE, L"b one"},
				{1, 2, flagVALUE, L"two"},
				{2, 3, flagVALUE, L"a three"},
				{1, 1, flagVALUE, L"one"},
			};

			// Do a stable sort on ulFlagName (first member)
			addin::SortFlagArray(flagArray, _countof(flagArray));

			testFA(L"0", flagArray[0], {1, 2, flagVALUE, L"two"});
			testFA(L"1", flagArray[1], {1, 1, flagVALUE, L"one"});
			testFA(L"2", flagArray[2], {2, 1, flagVALUE, L"b one"});
			testFA(L"3", flagArray[3], {2, 3, flagVALUE, L"a three"});

			auto flagV = std::vector<FLAG_ARRAY_ENTRY>{};
			addin::AppendFlagIfNotDupe(flagV, {1, 3, flagVALUE, L"three"});
			testFA(L"add 1", flagV[0], {1, 3, flagVALUE, L"three"});

			addin::AppendFlagIfNotDupe(flagV, {1, 4, flagVALUE, L"one"});
			testFA(L"add 2", flagV[1], {1, 4, flagVALUE, L"one"});
			addin::AppendFlagIfNotDupe(flagV, {1, 4, flagVALUE, L"one"});
			Assert::AreEqual(size_t{2}, flagV.size(), L"no dupe");

			addin::AppendFlagIfNotDupe(flagV, {2, 2, flagVALUE, L"c two"});
			testFA(L"add 3", flagV[2], {2, 2, flagVALUE, L"c two"});

			addin::MergeFlagArrays(flagV, flagArray, _countof(flagArray));
			Assert::AreEqual(size_t{7}, flagV.size(), L"merge size");

			testFA(L"m0", flagV[0], {1, 3, flagVALUE, L"three"});
			testFA(L"m1", flagV[1], {1, 4, flagVALUE, L"one"});
			testFA(L"m2", flagV[2], {1, 2, flagVALUE, L"two"});
			testFA(L"m3", flagV[3], {1, 1, flagVALUE, L"one"});
			testFA(L"m4", flagV[4], {2, 2, flagVALUE, L"c two"});
			testFA(L"m5", flagV[5], {2, 1, flagVALUE, L"b one"});
			testFA(L"m6", flagV[6], {2, 3, flagVALUE, L"a three"});

			auto flagV2 = std::vector<FLAG_ARRAY_ENTRY>{};
			addin::AppendFlagIfNotDupe(flagV2, {4, 3, flagVALUE, L"three"});
			testFA(L"f2 add 1", flagV2[0], {4, 3, flagVALUE, L"three"});
			addin::MergeFlagArrays(flagV2, nullptr, 0);
			Assert::AreEqual(size_t{1}, flagV2.size(), L"merge null");

			addin::AppendFlagIfNotDupe(flagV2, {4, 4, flagVALUE, L"four"});
			Assert::AreEqual(size_t{2}, flagV2.size(), L"merge null");
			testFA(L"f2 add 2", flagV2[1], {4, 4, flagVALUE, L"four"});

			FLAG_ARRAY_ENTRY flagArray2[] = {{3, 1, flagVALUE, L"four one"}};
			addin::MergeFlagArrays(flagV2, flagArray2, _countof(flagArray2));
			Assert::AreEqual(size_t{3}, flagV2.size(), L"merge null");
			testFA(L"f2 merge", flagV2[0], {3, 1, flagVALUE, L"four one"});
		}
	};
} // namespace addintest