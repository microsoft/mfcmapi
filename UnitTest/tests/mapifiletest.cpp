#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/mapiFile.h>
#include <core/utility/error.h>
#include <mapistub/library/stubutils.h>
#include <wrl/client.h>

namespace mapifiletest
{
	using Microsoft::WRL::ComPtr;
	const BYTE sentinel[]{0, 1, 2, 0xFF, 0, 0x80, 0x42, 7};

	struct CloseFile
	{
		void operator()(void* handle) const { CloseHandle(handle); }
	};
	using FileHandle = std::unique_ptr<void, CloseFile>;

	TEST_CLASS(mapifiletest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD_INITIALIZE(createStorage)
		{
			const auto initialized = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
			Assert::IsTrue(SUCCEEDED(initialized) || initialized == RPC_E_CHANGED_MODE);
			m_uninitialize = SUCCEEDED(initialized);
			m_mapiRequests = 0;
			m_originalMapiLog = std::move(mapistub::logLoadMapiCallback);
			mapistub::logLoadMapiCallback = [this](LPCWSTR, va_list) { ++m_mapiRequests; };
			InterlockedIncrement(reinterpret_cast<volatile LONG*>(&mapistub::g_ulDllSequenceNum));
			m_originalDisplayError = std::move(error::displayError);
			m_errors.clear();
			error::displayError = [this](const std::wstring& text) { m_errors.push_back(text); };
			wchar_t directory[MAX_PATH]{};
			wchar_t path[MAX_PATH]{};
			const auto length = GetCurrentDirectoryW(_countof(directory), directory);
			Assert::IsTrue(length != 0 && length < _countof(directory));
			Assert::IsTrue(GetTempFileNameW(directory, L"mfi", 0, path) != 0);
			m_path = path;
			m_attributes = GetFileAttributesW(path);
			Assert::IsTrue(m_attributes != INVALID_FILE_ATTRIBUTES);
			ComPtr<IStorage> storage;
			Assert::AreEqual<HRESULT>(
				S_OK,
				StgCreateDocfile(path, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, storage.GetAddressOf()));
			ComPtr<IStream> stream;
			Assert::AreEqual<HRESULT>(
				S_OK,
				storage->CreateStream(L"sentinel", STGM_WRITE | STGM_SHARE_EXCLUSIVE, 0, 0, stream.GetAddressOf()));
			ULONG written{};
			Assert::AreEqual<HRESULT>(S_OK, stream->Write(sentinel, sizeof(sentinel), &written));
			Assert::AreEqual<ULONG>(sizeof(sentinel), written);
			stream.Reset();
			Assert::AreEqual<HRESULT>(S_OK, storage->Commit(STGC_DEFAULT));
		}

		TEST_METHOD_CLEANUP(removeStorage)
		{
			error::displayError = std::move(m_originalDisplayError);
			mapistub::logLoadMapiCallback = std::move(m_originalMapiLog);
			const auto attributes = m_path.empty() || SetFileAttributesW(m_path.c_str(), m_attributes);
			const auto removed = m_path.empty() || DeleteFileW(m_path.c_str());
			if (m_uninitialize) CoUninitialize();
			Assert::IsTrue(attributes != FALSE, L"Could not restore the owned temporary file attributes.");
			Assert::IsTrue(removed != FALSE, L"Could not remove the owned temporary storage file.");
			Assert::AreEqual<ULONG>(0, m_mapiRequests, L"Storage tests must not resolve installed MAPI APIs.");
		}

		TEST_METHOD(UnlockedStorageKeepsReadWriteAccess) { CheckOpened(true, STGM_READWRITE); }

		TEST_METHOD(ReaderDenyingWritesAllowsReadOnlyFallback)
		{
			const auto lock = Lock(FILE_SHARE_READ);
			CheckOpened(true, STGM_READ);
		}

		TEST_METHOD(ReaderAllowingWritesKeepsReadWriteAccess)
		{
			const auto lock = Lock(FILE_SHARE_READ | FILE_SHARE_WRITE);
			CheckOpened(true, STGM_READWRITE);
		}

		TEST_METHOD(StorageReaderAllowsReadOnlyFallback)
		{
			ComPtr<IStorage> lock;
			Assert::AreEqual<HRESULT>(
				S_OK,
				StgOpenStorage(
					m_path.c_str(), nullptr, STGM_READ | STGM_SHARE_DENY_WRITE, nullptr, 0, lock.GetAddressOf()));
			CheckOpened(true, STGM_READ);
		}

		TEST_METHOD(ReadOnlyAttributeKeepsExistingFallback)
		{
			Assert::IsTrue(SetFileAttributesW(m_path.c_str(), m_attributes | FILE_ATTRIBUTE_READONLY) != FALSE);
			CheckOpened(true, STGM_READ);
		}

		TEST_METHOD(ExplicitReadOnlyDoesNotRequestWriteAccess)
		{
			const auto lock = Lock(FILE_SHARE_READ);
			CheckOpened(false, STGM_READ);
		}

		TEST_METHOD(ExclusiveLockPreservesSharingViolation)
		{
			const auto lock = Lock(0);
			ComPtr<IStorage> storage;
			Assert::AreEqual<HRESULT>(
				STG_E_SHAREVIOLATION, file::MyStgOpenStorage(m_path, true, storage.GetAddressOf()));
			Assert::IsNull(storage.Get());
		}

		TEST_METHOD(FailedMessageLoadReturnsStorageErrorWithoutResolvingMAPI)
		{
			const auto lock = Lock(0);
			auto result = S_OK;
			ComPtr<IMessage> message;
			message.Attach(file::LoadMSGToMessage(m_path, &result));
			Assert::IsNull(message.Get());
			Assert::AreEqual<HRESULT>(STG_E_SHAREVIOLATION, result);
			Assert::AreEqual<size_t>(1, m_errors.size());
		}

		TEST_METHOD(MissingMessageReturnsFileNotFoundWithoutResolvingMAPI)
		{
			const auto path = m_path + L".missing";
			Assert::IsTrue(GetFileAttributesW(path.c_str()) == INVALID_FILE_ATTRIBUTES);
			auto result = S_OK;
			ComPtr<IMessage> message;
			message.Attach(file::LoadMSGToMessage(path, &result));
			Assert::IsNull(message.Get());
			Assert::AreEqual<HRESULT>(STG_E_FILENOTFOUND, result);
			Assert::AreEqual<size_t>(1, m_errors.size());
		}

		TEST_METHOD(NullStorageOutputIsRejected)
		{
			Assert::AreEqual<HRESULT>(MAPI_E_INVALID_PARAMETER, file::MyStgOpenStorage(m_path, true, nullptr));
		}

		TEST_METHOD(MessageAttachmentNamesRequireAnExactExtension)
		{
			for (const auto name : {L"template.oft", L"template.OFT", L"message.msg", L"message.MsG"})
				Assert::IsTrue(file::IsMessageFileName(name));
			for (const auto name : {L"", L"template", L"template.oft.exe", L"file.doc", L"file.oft ", L"file.msgx"})
				Assert::IsFalse(file::IsMessageFileName(name));
			Assert::IsFalse(file::IsMessageAttachment(nullptr));
		}

		TEST_METHOD(AttachmentStorageOutlivesSourceAndIsAnIndependentCopy)
		{
			auto source = StorageStream();
			// OpenMSGStorage must rewind even if the provider returns a stream positioned at EOF.
			Assert::AreEqual<HRESULT>(S_OK, source->Seek({}, STREAM_SEEK_END, nullptr));
			ComPtr<IStorage> storage;
			Assert::AreEqual<HRESULT>(S_OK, file::OpenMSGStorage(source.Get(), storage.GetAddressOf()));
			source.Reset();
			ComPtr<IStream> contents;
			Assert::AreEqual<HRESULT>(
				S_OK,
				storage->OpenStream(
					L"sentinel", nullptr, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, contents.GetAddressOf()));
			BYTE actual[sizeof(sentinel)]{};
			ULONG read{};
			Assert::AreEqual<HRESULT>(S_OK, contents->Read(actual, sizeof(actual), &read));
			Assert::AreEqual<ULONG>(sizeof(sentinel), read);
			Assert::IsTrue(memcmp(actual, sentinel, sizeof(sentinel)) == 0);
			Assert::AreEqual<HRESULT>(S_OK, contents->Seek({}, STREAM_SEEK_SET, nullptr));
			const BYTE changed = 0xAA;
			Assert::AreEqual<HRESULT>(S_OK, contents->Write(&changed, sizeof(changed), nullptr));
			contents.Reset();
			Assert::AreEqual<HRESULT>(S_OK, storage->Commit(STGC_DEFAULT));
			storage.Reset();
			CheckOpened(false, STGM_READ);
		}

		TEST_METHOD(InvalidAttachmentStorageDoesNotResolveMAPI)
		{
			ComPtr<IStream> stream;
			Assert::AreEqual<HRESULT>(S_OK, CreateStreamOnHGlobal(nullptr, true, stream.GetAddressOf()));
			ComPtr<IStorage> storage;
			Assert::IsTrue(FAILED(file::OpenMSGStorage(stream.Get(), storage.GetAddressOf())));
			Assert::IsNull(storage.Get());
			Assert::AreEqual<HRESULT>(S_OK, stream->Write(sentinel, sizeof(sentinel), nullptr));
			Assert::IsTrue(FAILED(file::OpenMSGStorage(stream.Get(), storage.GetAddressOf())));
			Assert::IsNull(storage.Get());
		}

		TEST_METHOD(AttachmentStorageCopiesMultipleChunks)
		{
			ComPtr<ILockBytes> bytes;
			Assert::AreEqual<HRESULT>(S_OK, CreateILockBytesOnHGlobal(nullptr, true, bytes.GetAddressOf()));
			ComPtr<IStorage> original;
			Assert::AreEqual<HRESULT>(
				S_OK,
				StgCreateDocfileOnILockBytes(
					bytes.Get(), STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, original.GetAddressOf()));
			ComPtr<IStream> payload;
			Assert::AreEqual<HRESULT>(
				S_OK,
				original->CreateStream(L"large", STGM_WRITE | STGM_SHARE_EXCLUSIVE, 0, 0, payload.GetAddressOf()));
			std::vector<BYTE> expected(50000);
			for (size_t i = 0; i < expected.size(); ++i)
				expected[i] = static_cast<BYTE>(i);
			Assert::AreEqual<HRESULT>(
				S_OK, payload->Write(expected.data(), static_cast<ULONG>(expected.size()), nullptr));
			payload.Reset();
			Assert::AreEqual<HRESULT>(S_OK, original->Commit(STGC_DEFAULT));
			original.Reset();
			STATSTG stat{};
			Assert::AreEqual<HRESULT>(S_OK, bytes->Stat(&stat, STATFLAG_NONAME));
			std::vector<BYTE> raw(static_cast<size_t>(stat.cbSize.QuadPart));
			Assert::AreEqual<HRESULT>(S_OK, bytes->ReadAt({}, raw.data(), static_cast<ULONG>(raw.size()), nullptr));
			ComPtr<IStream> source;
			Assert::AreEqual<HRESULT>(S_OK, CreateStreamOnHGlobal(nullptr, true, source.GetAddressOf()));
			Assert::AreEqual<HRESULT>(S_OK, source->Write(raw.data(), static_cast<ULONG>(raw.size()), nullptr));
			ComPtr<IStorage> storage;
			Assert::AreEqual<HRESULT>(S_OK, file::OpenMSGStorage(source.Get(), storage.GetAddressOf()));
			source.Reset();
			bytes.Reset();
			Assert::AreEqual<HRESULT>(
				S_OK,
				storage->OpenStream(L"large", nullptr, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, payload.GetAddressOf()));
			std::vector<BYTE> actual(expected.size());
			ULONG read{};
			Assert::AreEqual<HRESULT>(S_OK, payload->Read(actual.data(), static_cast<ULONG>(actual.size()), &read));
			Assert::AreEqual<ULONG>(static_cast<ULONG>(actual.size()), read);
			Assert::IsTrue(actual == expected);
		}

		TEST_METHOD(TruncatedAttachmentStorageIsRejected)
		{
			auto stream = StorageStream();
			Assert::AreEqual<HRESULT>(S_OK, stream->SetSize(ULARGE_INTEGER{64}));
			ComPtr<IStorage> storage;
			Assert::IsTrue(FAILED(file::OpenMSGStorage(stream.Get(), storage.GetAddressOf())));
			Assert::IsNull(storage.Get());
		}

		TEST_METHOD(NullAttachmentStorageParametersAreRejected)
		{
			ComPtr<IStorage> storage;
			Assert::AreEqual<HRESULT>(MAPI_E_INVALID_PARAMETER, file::OpenMSGStorage(nullptr, storage.GetAddressOf()));
			Assert::IsNull(storage.Get());
			auto stream = StorageStream();
			Assert::AreEqual<HRESULT>(MAPI_E_INVALID_PARAMETER, file::OpenMSGStorage(stream.Get(), nullptr));
			auto result = S_OK;
			ComPtr<IMessage> message;
			message.Attach(file::LoadMSGToMessage(static_cast<LPSTORAGE>(nullptr), &result));
			Assert::IsNull(message.Get());
			Assert::AreEqual<HRESULT>(MAPI_E_INVALID_PARAMETER, result);
		}

	private:
		ComPtr<IStream> StorageStream()
		{
			const auto handle = Lock(FILE_SHARE_READ | FILE_SHARE_WRITE);
			LARGE_INTEGER size{};
			Assert::IsTrue(GetFileSizeEx(handle.get(), &size) != FALSE);
			Assert::IsTrue(size.QuadPart > 0 && size.QuadPart < MAXDWORD);
			std::vector<BYTE> bytes(static_cast<size_t>(size.QuadPart));
			DWORD read{};
			Assert::IsTrue(
				ReadFile(handle.get(), bytes.data(), static_cast<DWORD>(bytes.size()), &read, nullptr) != FALSE);
			Assert::AreEqual<DWORD>(static_cast<DWORD>(bytes.size()), read);
			ComPtr<IStream> stream;
			Assert::AreEqual<HRESULT>(S_OK, CreateStreamOnHGlobal(nullptr, true, stream.GetAddressOf()));
			Assert::AreEqual<HRESULT>(S_OK, stream->Write(bytes.data(), read, nullptr));
			return stream;
		}

		FileHandle Lock(DWORD share)
		{
			const auto handle = CreateFileW(m_path.c_str(), GENERIC_READ, share, nullptr, OPEN_EXISTING, 0, nullptr);
			Assert::IsTrue(handle != INVALID_HANDLE_VALUE);
			return FileHandle(handle);
		}

		void CheckOpened(bool bestAccess, DWORD access)
		{
			ComPtr<IStorage> storage;
			Assert::AreEqual<HRESULT>(S_OK, file::MyStgOpenStorage(m_path, bestAccess, storage.GetAddressOf()));
			STATSTG stat{};
			Assert::AreEqual<HRESULT>(S_OK, storage->Stat(&stat, STATFLAG_NONAME));
			Assert::AreEqual<DWORD>(access, stat.grfMode & (STGM_WRITE | STGM_READWRITE));
			ComPtr<IStream> stream;
			Assert::AreEqual<HRESULT>(
				S_OK,
				storage->OpenStream(L"sentinel", nullptr, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, stream.GetAddressOf()));
			BYTE actual[sizeof(sentinel)]{};
			ULONG read{};
			Assert::AreEqual<HRESULT>(S_OK, stream->Read(actual, sizeof(actual), &read));
			Assert::AreEqual<ULONG>(sizeof(sentinel), read);
			Assert::IsTrue(memcmp(actual, sentinel, sizeof(sentinel)) == 0);
		}

		std::wstring m_path;
		DWORD m_attributes{};
		bool m_uninitialize{};
		ULONG m_mapiRequests{};
		std::function<void(LPCWSTR, va_list)> m_originalMapiLog;
		std::function<void(const std::wstring&)> m_originalDisplayError;
		std::vector<std::wstring> m_errors;
	};
} // namespace mapifiletest
