#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/compressedRtf.h>
#include <core/mapi/mapiFile.h>
#include <memory>
#include <wrl/client.h>

#ifdef _M_IX86
#pragma comment(linker, "/export:MAPIInitialize@4=_TestRtfMAPIInitialize@4")
#pragma comment(linker, "/export:MAPIUninitialize@0=_TestRtfMAPIUninitialize@0")
#pragma comment(linker, "/export:WrapCompressedRTFStream@12=_TestRtfWrapCompressedRTFStream@12")
#pragma comment(linker, "/export:WrapCompressedRTFStreamEx@16=_TestRtfWrapCompressedRTFStreamEx@16")
#else
#pragma comment(linker, "/export:MAPIInitialize=TestRtfMAPIInitialize")
#pragma comment(linker, "/export:MAPIUninitialize=TestRtfMAPIUninitialize")
#pragma comment(linker, "/export:WrapCompressedRTFStream=TestRtfWrapCompressedRTFStream")
#pragma comment(linker, "/export:WrapCompressedRTFStreamEx=TestRtfWrapCompressedRTFStreamEx")
#endif

namespace
{
	HRESULT initializeResult{S_OK};
	HRESULT wrapResult{S_OK};
	ULONG initializeCalls{};
	ULONG uninitializeCalls{};
	ULONG wrapCalls{};
	ULONG wrapFlags{};
	ULONG wrapInCodePage{};
	ULONG wrapOutCodePage{};
	ULONG wrapStreamFlags{};
	bool nullWrappedStream{};
	std::vector<BYTE> wrappedInput;
	std::vector<BYTE> wrappedOutput;

	class TestRtfStream final : public IStream
	{
	public:
		TestRtfStream() { Assert::AreEqual(S_OK, CreateStreamOnHGlobal(nullptr, true, &m_stream)); }
		~TestRtfStream()
		{
			m_stream->Release();
			if (target) target->Release();
		}
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void** object) override
		{
			*object = nullptr;
			return E_NOINTERFACE;
		}
		ULONG STDMETHODCALLTYPE AddRef() override { return ++m_refs; }
		ULONG STDMETHODCALLTYPE Release() override
		{
			const auto refs = --m_refs;
			if (!refs) delete this;
			return refs;
		}
		HRESULT STDMETHODCALLTYPE Read(void* buffer, ULONG count, ULONG* read) override
		{
			if (FAILED(readResult)) return readResult;
			return m_stream->Read(buffer, count, read);
		}
		HRESULT STDMETHODCALLTYPE Write(const void* buffer, ULONG count, ULONG* written) override
		{
			if (FAILED(writeResult)) return writeResult;
			if (shortWrite && count) --count;
			return m_stream->Write(buffer, count, written);
		}
		HRESULT STDMETHODCALLTYPE Seek(LARGE_INTEGER move, DWORD origin, ULARGE_INTEGER* position) override
		{
			return m_stream->Seek(move, origin, position);
		}
		HRESULT STDMETHODCALLTYPE SetSize(ULARGE_INTEGER size) override
		{
			if (FAILED(sizeResult)) return sizeResult;
			return m_stream->SetSize(size);
		}
		HRESULT STDMETHODCALLTYPE CopyTo(IStream*, ULARGE_INTEGER, ULARGE_INTEGER*, ULARGE_INTEGER*) override
		{
			return E_NOTIMPL;
		}
		HRESULT STDMETHODCALLTYPE Commit(DWORD flags) override
		{
			++commits;
			if (FAILED(commitResult)) return commitResult;
			if (target)
			{
				STATSTG stat{};
				auto hRes = m_stream->Stat(&stat, STATFLAG_NONAME);
				if (FAILED(hRes)) return hRes;
				wrappedInput.resize(static_cast<size_t>(stat.cbSize.QuadPart));
				hRes = m_stream->Seek({}, STREAM_SEEK_SET, nullptr);
				if (SUCCEEDED(hRes))
					hRes = m_stream->Read(wrappedInput.data(), static_cast<ULONG>(wrappedInput.size()), nullptr);
				if (SUCCEEDED(hRes))
					hRes = target->Write(wrappedOutput.data(), static_cast<ULONG>(wrappedOutput.size()), nullptr);
				return hRes;
			}

			return m_stream->Commit(flags);
		}
		HRESULT STDMETHODCALLTYPE Revert() override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE LockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE UnlockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE Stat(STATSTG* stat, DWORD flags) override
		{
			const auto hRes = m_stream->Stat(stat, flags);
			if (SUCCEEDED(hRes) && advertisedSize) stat->cbSize.QuadPart = advertisedSize;
			return hRes;
		}
		HRESULT STDMETHODCALLTYPE Clone(IStream**) override { return E_NOTIMPL; }

		LPSTREAM target{};
		HRESULT readResult{S_OK};
		HRESULT writeResult{S_OK};
		HRESULT sizeResult{S_OK};
		HRESULT commitResult{S_OK};
		bool shortWrite{};
		ULONG commits{};
		ULONGLONG advertisedSize{};

	private:
		ULONG m_refs{1};
		LPSTREAM m_stream{};
	};

	HRESULT compressionWriteResult{S_OK};
	HRESULT compressionCommitResult{S_OK};
	bool compressionShortWrite{};
} // namespace

extern "C" HRESULT STDAPICALLTYPE TestRtfMAPIInitialize(LPVOID)
{
	++initializeCalls;
	return initializeResult;
}

extern "C" void STDAPICALLTYPE TestRtfMAPIUninitialize() { ++uninitializeCalls; }

extern "C" HRESULT STDAPICALLTYPE TestRtfWrapCompressedRTFStream(LPSTREAM source, ULONG flags, LPSTREAM* result)
{
	++wrapCalls;
	wrapFlags = flags;
	*result = nullptr;
	BYTE bytes[128]{};
	ULONG read{};
	const auto hRes = source->Read(bytes, sizeof(bytes), &read);
	if (FAILED(hRes)) return hRes;
	wrappedInput.assign(bytes, bytes + read);
	if (FAILED(wrapResult)) return wrapResult;
	if (nullWrappedStream) return S_OK;
	if (flags & MAPI_MODIFY)
	{
		auto stream = new TestRtfStream();
		stream->target = source;
		source->AddRef();
		stream->writeResult = compressionWriteResult;
		stream->commitResult = compressionCommitResult;
		stream->shortWrite = compressionShortWrite;
		*result = stream;
		return S_OK;
	}

	auto streamResult = CreateStreamOnHGlobal(nullptr, true, result);
	if (FAILED(streamResult)) return streamResult;
	streamResult = (*result)->Write(wrappedOutput.data(), static_cast<ULONG>(wrappedOutput.size()), nullptr);
	if (SUCCEEDED(streamResult)) streamResult = (*result)->Seek({}, STREAM_SEEK_SET, nullptr);
	return streamResult;
}

extern "C" HRESULT STDAPICALLTYPE
TestRtfWrapCompressedRTFStreamEx(LPSTREAM source, const RTF_WCSINFO* info, LPSTREAM* result, RTF_WCSRETINFO* returned)
{
	Assert::AreEqual(static_cast<ULONG>(sizeof(RTF_WCSINFO)), info->size);
	Assert::AreEqual(static_cast<ULONG>(sizeof(RTF_WCSRETINFO)), returned->size);
	wrapInCodePage = info->ulInCodePage;
	wrapOutCodePage = info->ulOutCodePage;
	returned->ulStreamFlags = wrapStreamFlags;
	return TestRtfWrapCompressedRTFStream(source, info->ulFlags, result);
}

namespace compressedrtftest
{
	const std::wstring mela = L"180000000c0000004d454c41000000007b5c7274663120746573747d";
	// Literal LZFu tokens followed by the end reference to dictionary offset 207 + 12.
	const std::wstring lzfu = L"1c0000000c0000004c5a4675236427fa007b5c727466312074106573747d0db0";

	class TestRtfProperty final : public IMAPIProp
	{
	public:
		TestRtfProperty() : stream(new TestRtfStream()) {}
		~TestRtfProperty() { stream->Release(); }
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void** object) override
		{
			*object = nullptr;
			return E_NOINTERFACE;
		}
		ULONG STDMETHODCALLTYPE AddRef() override { return ++m_refs; }
		ULONG STDMETHODCALLTYPE Release() override { return --m_refs; }
		HRESULT STDMETHODCALLTYPE GetLastError(HRESULT, ULONG, LPMAPIERROR*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SaveChanges(ULONG flags) override
		{
			++saves;
			saveFlags = flags;
			Assert::AreEqual(1UL, stream->commits);
			return saveResult;
		}
		HRESULT STDMETHODCALLTYPE GetProps(LPSPropTagArray, ULONG, ULONG*, LPSPropValue*) override
		{
			++getPropsCalls;
			return MAPI_E_NOT_ENOUGH_MEMORY;
		}
		HRESULT STDMETHODCALLTYPE GetPropList(ULONG, LPSPropTagArray*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE
		OpenProperty(ULONG tag, LPCIID iid, ULONG mode, ULONG flags, LPUNKNOWN* object) override
		{
			++opens;
			Assert::AreEqual(static_cast<ULONG>(PR_RTF_COMPRESSED), tag);
			Assert::IsTrue(*iid == IID_IStream);
			openMode = mode;
			openFlags = flags;
			*object = nullptr;
			if (FAILED(openResult) || nullStream) return openResult;
			const auto hRes = stream->Seek({}, STREAM_SEEK_SET, nullptr);
			if (FAILED(hRes)) return hRes;
			stream->AddRef();
			*object = stream;
			return S_OK;
		}
		HRESULT STDMETHODCALLTYPE SetProps(ULONG, LPSPropValue, LPSPropProblemArray*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE DeleteProps(LPSPropTagArray, LPSPropProblemArray*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE
		CopyTo(ULONG, LPCIID, LPSPropTagArray, ULONG_PTR, LPMAPIPROGRESS, LPCIID, LPVOID, ULONG, LPSPropProblemArray*)
			override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE
		CopyProps(LPSPropTagArray, ULONG_PTR, LPMAPIPROGRESS, LPCIID, LPVOID, ULONG, LPSPropProblemArray*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE GetNamesFromIDs(LPSPropTagArray*, LPGUID, ULONG, ULONG*, LPMAPINAMEID**) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE GetIDsFromNames(ULONG, LPMAPINAMEID*, ULONG, LPSPropTagArray*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		TestRtfStream* stream;
		HRESULT openResult{S_OK};
		HRESULT saveResult{S_OK};
		ULONG opens{};
		ULONG saves{};
		ULONG getPropsCalls{};
		ULONG openMode{};
		ULONG openFlags{};
		ULONG saveFlags{};
		bool nullStream{};

	private:
		ULONG m_refs{1};
	};

	void RequireMapi()
	{
		const auto initialized = MAPIInitialize(nullptr);
		Assert::IsTrue(
			SUCCEEDED(initialized), L"Real decoder integration requires an installed matching-bitness MAPI runtime.");
		MAPIUninitialize();
	}

	TEST_CLASS(compressedrtfintegrationtest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }
		TEST_METHOD_INITIALIZE(setup)
		{
			const auto hRes = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
			Assert::IsTrue(SUCCEEDED(hRes) || hRes == RPC_E_CHANGED_MODE);
			m_uninitialize = SUCCEEDED(hRes);
			const auto initialized = MAPIInitialize(nullptr);
			m_mapiInitialized = SUCCEEDED(initialized);
			Assert::IsTrue(m_mapiInitialized, L"These tests require matching-bitness installed MAPI.");
		}
		TEST_METHOD_CLEANUP(cleanup)
		{
			const auto removed = m_path.empty() || DeleteFileW(m_path.c_str());
			if (m_mapiInitialized) MAPIUninitialize();
			if (m_uninitialize) CoUninitialize();
			Assert::IsTrue(removed != FALSE, L"Could not remove the owned temporary MSG file.");
		}

		TEST_METHOD(DecompressesMelaAndFormattedHex)
		{
			RequireMapi();
			std::vector<BYTE> output;
			Assert::AreEqual(
				S_OK,
				mapi::DecompressRtfHex(L" 0x18 000000\r\n0c000000 4d454c41 00000000 7b5c7274663120746573747d", output));
			Assert::AreEqual(strings::HexStringToBin(L"7b5c7274663120746573747d"), output);
		}

		TEST_METHOD(DecompressesLzfu)
		{
			RequireMapi();
			std::vector<BYTE> output;
			Assert::AreEqual(S_OK, mapi::DecompressRtfHex(lzfu, output));
			Assert::AreEqual(strings::HexStringToBin(L"7b5c7274663120746573747d"), output);
		}

		TEST_METHOD(PreservesEmbeddedNullAndHighBytes)
		{
			RequireMapi();
			std::vector<BYTE> output;
			Assert::AreEqual(S_OK, mapi::DecompressRtfHex(L"10000000040000004d454c410000000000ff0080", output));
			Assert::AreEqual(std::vector<BYTE>({0, 0xff, 0, 0x80}), output);
		}

		TEST_METHOD(RejectsDecodedLengthMismatch)
		{
			RequireMapi();
			auto hex = lzfu;
			hex[9] = L'd';
			std::vector<BYTE> output{1};
			Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::DecompressRtfHex(hex, output));
			Assert::IsTrue(output.empty());
		}

		TEST_METHOD(AcceptsEmptyMela)
		{
			RequireMapi();
			std::vector<BYTE> output{1};
			Assert::AreEqual(S_OK, mapi::DecompressRtfHex(L"0c000000000000004d454c4100000000", output));
			Assert::IsTrue(output.empty());
		}

		TEST_METHOD(CompressesAndRoundTripsExactBytes)
		{
			RequireMapi();
			for (const auto& input :
				 {strings::HexStringToBin(L"7b5c7274663120746573747d"),
				  std::vector<BYTE>{0, 0xff, 0, 0x80},
				  std::vector<BYTE>{},
				  std::vector<BYTE>(65536, 'a')})
			{
				std::vector<BYTE> compressed{1};
				Assert::AreEqual(S_OK, mapi::CompressRtf(input, compressed));
				Assert::IsTrue(compressed.size() >= 16);
				std::vector<BYTE> output{1};
				Assert::AreEqual(
					S_OK,
					mapi::DecompressRtfHex(strings::BinToHexString(compressed, false), output),
					strings::format(
						L"Input bytes: %zu; header: %ws",
						input.size(),
						strings::BinToHexString(compressed.data(), min(size_t{16}, compressed.size()), false).c_str())
						.c_str());
				Assert::AreEqual(input, output);
			}
		}

		TEST_METHOD(CompressesAndRoundTripsBeyondFormerRawSizeLimit)
		{
			RequireMapi();
			const std::vector<BYTE> input(17 * 1024 * 1024, 'a');
			std::vector<BYTE> compressed;
			Assert::AreEqual(S_OK, mapi::CompressRtf(input, compressed));
			std::vector<BYTE> output;
			Assert::AreEqual(S_OK, mapi::DecompressRtfHex(strings::BinToHexString(compressed, false), output));
			Assert::AreEqual(input, output);
		}

		TEST_METHOD(CompressesBeyondFormerCompressedSizeLimit)
		{
			RequireMapi();
			std::vector<BYTE> input(4 * 1024 * 1024);
			ULONG state = 1;
			for (auto& byte : input)
			{
				state ^= state << 13;
				state ^= state >> 17;
				state ^= state << 5;
				byte = static_cast<BYTE>(state);
			}

			std::vector<BYTE> compressed{1};
			Assert::AreEqual(S_OK, mapi::CompressRtf(input, compressed));
			Assert::IsTrue(compressed.size() > 4 * 1024 * 1024);
			std::vector<BYTE> output;
			Assert::AreEqual(S_OK, mapi::DecompressRtf(compressed, output));
			Assert::AreEqual(input, output);
		}

		TEST_METHOD(SavesMessageAndReadsExactBytesAfterReopening)
		{
			wchar_t directory[MAX_PATH]{};
			wchar_t path[MAX_PATH]{};
			const auto length = GetTempPathW(_countof(directory), directory);
			Assert::IsTrue(length != 0 && length < _countof(directory));
			Assert::IsTrue(GetTempFileNameW(directory, L"rtf", 0, path) != 0);
			m_path = path;
			Microsoft::WRL::ComPtr<IMessage> message;
			Microsoft::WRL::ComPtr<IStorage> storage;
			Assert::AreEqual(S_OK, file::CreateNewMSG(m_path, true, message.GetAddressOf(), storage.GetAddressOf()));
			const auto input = strings::HexStringToBin(L"7b5c72746631207065727369737465647d");
			std::vector<BYTE> compressed;
			Assert::AreEqual(S_OK, mapi::CompressRtf(input, compressed));
			Assert::AreEqual(S_OK, mapi::SaveCompressedRtf(message.Get(), compressed));
			Assert::AreEqual(S_OK, storage->Commit(STGC_DEFAULT));
			message.Reset();
			storage.Reset();
			auto hRes = S_OK;
			message.Attach(file::LoadMSGToMessage(m_path, &hRes));
			Assert::AreEqual(S_OK, hRes);
			Assert::IsNotNull(message.Get());
			std::vector<BYTE> reopened;
			Assert::AreEqual(S_OK, mapi::ReadCompressedRtf(message.Get(), reopened));
			Assert::AreEqual(compressed, reopened);
			std::vector<BYTE> output;
			Assert::AreEqual(S_OK, mapi::DecompressRtfHex(strings::BinToHexString(reopened, false), output));
			Assert::AreEqual(input, output);
		}

	private:
		bool m_uninitialize{};
		bool m_mapiInitialized{};
		std::wstring m_path;
	};

	TEST_CLASS(compressedrtftest){public : TEST_CLASS_INITIALIZE(initialize){unittest::init();
} // namespace compressedrtftest
TEST_METHOD(RejectsEmptyOddAndNonHexInputAndClearsOutput)
{
	for (const auto& hex : {L"", L" \r\n", L"0x", L"123", L"zz", L"00 gg"})
	{
		std::vector<BYTE> output{1, 2, 3};
		Assert::AreEqual(E_INVALIDARG, mapi::DecompressRtfHex(hex, output));
		Assert::IsTrue(output.empty());
	}
}

TEST_METHOD(RejectsTruncatedHeaderAndPayload)
{
	for (const auto& hex : {std::wstring(L"0102"), mela.substr(0, mela.size() - 2), mela + L"00"})
	{
		std::vector<BYTE> output{1};
		Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::DecompressRtfHex(hex, output));
		Assert::IsTrue(output.empty());
	}
}

TEST_METHOD(RejectsBadMagicSizeAndChecksum)
{
	for (const auto offset : {0, 8, 16, 24})
	{
		auto hex = mela;
		hex[offset] = L'f';
		std::vector<BYTE> output{1};
		Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::DecompressRtfHex(hex, output));
		Assert::IsTrue(output.empty());
	}

	auto corrupt = lzfu;
	corrupt.back() = L'1';
	std::vector<BYTE> output{1};
	Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::DecompressRtfHex(corrupt, output));
	Assert::IsTrue(output.empty());
}
}
;
TEST_CLASS(compressedrtfstreamtest)
{
public:
	TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }
	TEST_METHOD_INITIALIZE(setup)
	{
		m_memory = std::make_unique<unittest::testMapiMemory>();
		initializeResult = S_OK;
		wrapResult = S_OK;
		initializeCalls = uninitializeCalls = wrapCalls = 0;
		wrapFlags = ULONG_MAX;
		wrapInCodePage = wrapOutCodePage = wrapStreamFlags = 0;
		nullWrappedStream = false;
		compressionWriteResult = compressionCommitResult = S_OK;
		compressionShortWrite = false;
		wrappedInput.clear();
		wrappedOutput = strings::HexStringToBin(L"7b5c7274663120746573747d");
	}
	TEST_METHOD_CLEANUP(cleanup)
	{
		m_memory.reset();
		Assert::AreEqual(0UL, initializeCalls, L"Conversion must not own MAPI initialization.");
		Assert::AreEqual(0UL, uninitializeCalls, L"Conversion must not uninitialize its caller's MAPI.");
	}

	TEST_METHOD(UsesReadOnlyWrapperWithRewoundInput)
	{
		std::vector<BYTE> output;
		Assert::AreEqual(S_OK, mapi::DecompressRtfHex(lzfu, output));
		Assert::AreEqual(wrappedOutput, output);
		Assert::AreEqual(strings::HexStringToBin(lzfu), wrappedInput);
		Assert::AreEqual(0UL, wrapFlags);
		Assert::AreEqual(1UL, wrapCalls);
	}

	TEST_METHOD(ReportsWrapperFailure)
	{
		wrapResult = E_FAIL;
		std::vector<BYTE> output{1};
		Assert::AreEqual(wrapResult, mapi::DecompressRtfHex(mela, output));
		Assert::IsTrue(output.empty());
	}

	TEST_METHOD(RejectsShortAndLongDecodedStreams)
	{
		for (const auto length : {11, 13})
		{
			wrappedOutput.resize(length);
			std::vector<BYTE> output{1};
			Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::DecompressRtfHex(mela, output));
			Assert::IsTrue(output.empty());
		}
	}

	TEST_METHOD(WrapExPassesOptionsAndAllowsNativeBodyWithDifferentLength)
	{
		wrappedOutput = {0x41, 0, 0x42, 0};
		wrapStreamFlags = MAPI_NATIVE_BODY_TYPE_PLAINTEXT;
		std::vector<BYTE> output;
		ULONG streamFlags{};
		Assert::AreEqual(
			S_OK,
			mapi::DecompressRtf(
				strings::HexStringToBin(mela), output, true, MAPI_NATIVE_BODY, 1252, 1200, &streamFlags));
		Assert::AreEqual(wrappedOutput, output);
		Assert::AreEqual(static_cast<ULONG>(MAPI_NATIVE_BODY), wrapFlags);
		Assert::AreEqual(1252UL, wrapInCodePage);
		Assert::AreEqual(1200UL, wrapOutCodePage);
		Assert::AreEqual(wrapStreamFlags, streamFlags);
	}

	TEST_METHOD(WrapExReportsWrapperFailureAndClearsOutput)
	{
		wrapResult = E_FAIL;
		std::vector<BYTE> output{1};
		Assert::AreEqual(E_FAIL, mapi::DecompressRtf(strings::HexStringToBin(mela), output, true, MAPI_NATIVE_BODY));
		Assert::IsTrue(output.empty());
	}

	TEST_METHOD(RejectsHugeAdvertisedRawSizeWithoutAllocatingIt)
	{
		auto bytes = strings::HexStringToBin(lzfu);
		const ULONG rawSize = ULONG_MAX;
		memcpy(bytes.data() + sizeof(ULONG), &rawSize, sizeof(rawSize));
		std::vector<BYTE> output{1};
		Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::DecompressRtf(bytes, output));
		Assert::IsTrue(output.empty());
	}
	TEST_METHOD(RejectsSuccessfulWrapperWithoutOutputStream)
	{
		nullWrappedStream = true;
		std::vector<BYTE> output{1};
		Assert::AreEqual(E_UNEXPECTED, mapi::DecompressRtfHex(mela, output));
		Assert::IsTrue(output.empty());
	}

	TEST_METHOD(InvalidDataDoesNotInitializeOrWrap)
	{
		std::vector<BYTE> output{1};
		Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::DecompressRtfHex(L"0102", output));
		Assert::IsTrue(output.empty());
		Assert::AreEqual(0UL, initializeCalls);
		Assert::AreEqual(0UL, wrapCalls);
	}

	TEST_METHOD(CompressionReportsWrapperFailure)
	{
		std::vector<BYTE> output{1};
		wrapResult = E_FAIL;
		Assert::AreEqual(wrapResult, mapi::CompressRtf({1}, output));
		Assert::IsTrue(output.empty());
		Assert::AreEqual(static_cast<ULONG>(MAPI_MODIFY), wrapFlags);
	}

	TEST_METHOD(CompressionRejectsSuccessfulWrapperWithoutStream)
	{
		nullWrappedStream = true;
		std::vector<BYTE> output{1};
		Assert::AreEqual(E_UNEXPECTED, mapi::CompressRtf({}, output));
		Assert::IsTrue(output.empty());
	}

	TEST_METHOD(CompressionWritesAndCommitsExactBytes)
	{
		const auto input = wrappedOutput;
		wrappedOutput = strings::HexStringToBin(mela);
		std::vector<BYTE> output;
		Assert::AreEqual(S_OK, mapi::CompressRtf(input, output));
		Assert::AreEqual(input, wrappedInput);
		Assert::AreEqual(wrappedOutput, output);
		Assert::AreEqual(static_cast<ULONG>(MAPI_MODIFY), wrapFlags);
	}

	TEST_METHOD(CompressionReportsWriteAndCommitFailures)
	{
		const std::vector<BYTE> input{1};
		std::vector<BYTE> output{1};
		compressionWriteResult = STG_E_ACCESSDENIED;
		Assert::AreEqual(compressionWriteResult, mapi::CompressRtf(input, output));
		Assert::IsTrue(output.empty());
		compressionWriteResult = S_OK;
		compressionShortWrite = true;
		Assert::AreEqual(STG_E_WRITEFAULT, mapi::CompressRtf(input, output));
		Assert::IsTrue(output.empty());
		compressionShortWrite = false;
		compressionCommitResult = STG_E_MEDIUMFULL;
		Assert::AreEqual(compressionCommitResult, mapi::CompressRtf(input, output));
		Assert::IsTrue(output.empty());
	}

	TEST_METHOD(CompressionRejectsMalformedAndWrongSizedOutput)
	{
		std::vector<BYTE> output{1};
		wrappedOutput = {1};
		Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::CompressRtf({}, output));
		Assert::IsTrue(output.empty());
		wrappedOutput = strings::HexStringToBin(mela);
		Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::CompressRtf({}, output));
		Assert::IsTrue(output.empty());
	}

	TEST_METHOD(ReadsPropertyStreamWithoutGetPropsEvenWhenGetPropsWouldFail)
	{
		TestRtfProperty prop;
		const auto bytes = strings::HexStringToBin(mela);
		Assert::AreEqual(S_OK, prop.stream->Write(bytes.data(), static_cast<ULONG>(bytes.size()), nullptr));
		std::vector<BYTE> output{1};
		Assert::AreEqual(S_OK, mapi::ReadCompressedRtf(&prop, output));
		Assert::AreEqual(bytes, output);
		Assert::AreEqual(static_cast<ULONG>(STGM_READ), prop.openMode);
		Assert::AreEqual(0UL, prop.openFlags);
		Assert::AreEqual(0UL, prop.stream->commits);
		Assert::AreEqual(0UL, prop.saves);
		Assert::AreEqual(0UL, prop.getPropsCalls);
	}

	TEST_METHOD(SavesExactCompressedBytesAndTruncatesOldStream)
	{
		TestRtfProperty prop;
		auto bytes = strings::HexStringToBin(mela);
		bytes.resize(1000);
		Assert::AreEqual(S_OK, prop.stream->Write(bytes.data(), static_cast<ULONG>(bytes.size()), nullptr));
		bytes = strings::HexStringToBin(mela);
		Assert::AreEqual(S_OK, mapi::SaveCompressedRtf(&prop, bytes));
		Assert::AreEqual(static_cast<ULONG>(STGM_READWRITE), prop.openMode);
		Assert::AreEqual(static_cast<ULONG>(MAPI_CREATE | MAPI_MODIFY), prop.openFlags);
		Assert::AreEqual(1UL, prop.saves);
		Assert::AreEqual(static_cast<ULONG>(KEEP_OPEN_READWRITE), prop.saveFlags);
		std::vector<BYTE> output;
		Assert::AreEqual(S_OK, mapi::ReadCompressedRtf(&prop, output));
		Assert::AreEqual(bytes, output);
	}

	TEST_METHOD(InvalidMessageDataNeverOpensForWrite)
	{
		TestRtfProperty prop;
		Assert::AreEqual(MAPI_E_CORRUPT_DATA, mapi::SaveCompressedRtf(&prop, {1}));
		Assert::AreEqual(0UL, prop.opens);
		Assert::AreEqual(0UL, prop.saves);
	}

	TEST_METHOD(MessageOpenAndReadFailuresAreReturned)
	{
		TestRtfProperty prop;
		std::vector<BYTE> output{1};
		prop.openResult = MAPI_E_NOT_FOUND;
		Assert::AreEqual(prop.openResult, mapi::ReadCompressedRtf(&prop, output));
		Assert::IsTrue(output.empty());
		Assert::AreEqual(prop.openResult, mapi::SaveCompressedRtf(&prop, strings::HexStringToBin(mela)));
		prop.openResult = S_OK;
		prop.nullStream = true;
		Assert::AreEqual(E_UNEXPECTED, mapi::ReadCompressedRtf(&prop, output));
		Assert::AreEqual(E_UNEXPECTED, mapi::SaveCompressedRtf(&prop, strings::HexStringToBin(mela)));
		prop.nullStream = false;
		prop.stream->readResult = STG_E_READFAULT;
		Assert::AreEqual(STG_E_READFAULT, mapi::ReadCompressedRtf(&prop, output));
		Assert::IsTrue(output.empty());
		Assert::AreEqual(0UL, prop.saves);
	}

	TEST_METHOD(MessageWriteSizeAndCommitFailuresNeverSave)
	{
		TestRtfProperty prop;
		const auto bytes = strings::HexStringToBin(mela);
		prop.stream->writeResult = STG_E_ACCESSDENIED;
		Assert::AreEqual(STG_E_ACCESSDENIED, mapi::SaveCompressedRtf(&prop, bytes));
		prop.stream->writeResult = S_OK;
		prop.stream->shortWrite = true;
		Assert::AreEqual(STG_E_WRITEFAULT, mapi::SaveCompressedRtf(&prop, bytes));
		prop.stream->shortWrite = false;
		prop.stream->sizeResult = STG_E_MEDIUMFULL;
		Assert::AreEqual(STG_E_MEDIUMFULL, mapi::SaveCompressedRtf(&prop, bytes));
		prop.stream->sizeResult = S_OK;
		prop.stream->commitResult = E_FAIL;
		Assert::AreEqual(E_FAIL, mapi::SaveCompressedRtf(&prop, bytes));
		Assert::AreEqual(0UL, prop.saves);
	}

	TEST_METHOD(MessageSaveFailureIsReturned)
	{
		TestRtfProperty prop;
		prop.saveResult = MAPI_E_NO_ACCESS;
		Assert::AreEqual(prop.saveResult, mapi::SaveCompressedRtf(&prop, strings::HexStringToBin(mela)));
		Assert::AreEqual(1UL, prop.saves);
	}

	TEST_METHOD(MessageReadRejectsOversizedProperty)
	{
		TestRtfProperty prop;
		ULARGE_INTEGER size{};
		size.QuadPart = static_cast<ULONGLONG>(ULONG_MAX) + 1;
		prop.stream->advertisedSize = size.QuadPart;
		std::vector<BYTE> output{1};
		Assert::AreEqual(MAPI_E_TOO_BIG, mapi::ReadCompressedRtf(&prop, output));
		Assert::IsTrue(output.empty());
	}

private:
	std::unique_ptr<unittest::testMapiMemory> m_memory;
};
} // namespace compressedrtftest
