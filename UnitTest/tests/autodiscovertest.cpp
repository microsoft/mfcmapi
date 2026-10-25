#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/autoDiscover.h>
#include <core/utility/error.h>
#include <wrl/client.h>

#ifdef _M_IX86
#pragma comment(linker, "/export:HrGetAutoDiscoverXML@20=_TestHrGetAutoDiscoverXML@20")
#else
#pragma comment(linker, "/export:HrGetAutoDiscoverXML=TestHrGetAutoDiscoverXML")
#endif

namespace
{
	IStream* response{};
	HRESULT responseResult{S_OK};
	ULONG requests{};
	std::wstring requestedAddress;
} // namespace

extern "C" HRESULT STDAPICALLTYPE
TestHrGetAutoDiscoverXML(LPCWSTR address, LPCWSTR password, HANDLE cancel, ULONG flags, IStream** stream)
{
	Assert::IsNull(password);
	Assert::IsNull(cancel);
	Assert::AreEqual(0UL, flags);
	++requests;
	requestedAddress = address;
	*stream = response;
	if (response) response->AddRef();
	return responseResult;
}

namespace autodiscovertest
{
	class TestStream final : public IStream
	{
	public:
		explicit TestStream(const std::vector<BYTE>& bytes)
		{
			Assert::AreEqual(S_OK, CreateStreamOnHGlobal(nullptr, true, m_stream.GetAddressOf()));
			Assert::AreEqual(S_OK, m_stream->Write(bytes.data(), static_cast<ULONG>(bytes.size()), nullptr));
		}

		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void** object) override
		{
			*object = nullptr;
			return E_NOINTERFACE;
		}
		ULONG STDMETHODCALLTYPE AddRef() override { return ++refs; }
		ULONG STDMETHODCALLTYPE Release() override { return --refs; }
		HRESULT STDMETHODCALLTYPE Read(void* buffer, ULONG count, ULONG* read) override
		{
			++reads;
			if (FAILED(readResult))
			{
				*read = 0;
				return readResult;
			}

			const auto result = m_stream->Read(buffer, count, read);
			return partialAsFalse && result == S_OK && *read < count ? S_FALSE : result;
		}
		HRESULT STDMETHODCALLTYPE Seek(LARGE_INTEGER offset, DWORD origin, ULARGE_INTEGER* position) override
		{
			++seeks;
			if (FAILED(seekResult)) return seekResult;
			return m_stream->Seek(offset, origin, position);
		}
		HRESULT STDMETHODCALLTYPE Write(const void*, ULONG, ULONG*) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE SetSize(ULARGE_INTEGER) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE CopyTo(IStream*, ULARGE_INTEGER, ULARGE_INTEGER*, ULARGE_INTEGER*) override
		{
			return E_NOTIMPL;
		}
		HRESULT STDMETHODCALLTYPE Commit(DWORD) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE Revert() override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE LockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE UnlockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE Stat(STATSTG*, DWORD) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE Clone(IStream**) override { return E_NOTIMPL; }

		ULONG refs{1}, reads{}, seeks{};
		HRESULT readResult{S_OK}, seekResult{S_OK};
		bool partialAsFalse{};

	private:
		Microsoft::WRL::ComPtr<IStream> m_stream;
	};

	TEST_CLASS(autodiscovertest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD_INITIALIZE(captureErrors)
		{
			m_memory = std::make_unique<unittest::testMapiMemory>();
			m_originalError = std::move(error::displayError);
			error::displayError = [this](const auto& text) { m_errors.push_back(text); };
			m_errors.clear();
			response = nullptr;
			responseResult = S_OK;
			requests = 0;
		}

		TEST_METHOD_CLEANUP(restoreErrors)
		{
			response = nullptr;
			error::displayError = std::move(m_originalError);
			m_memory.reset();
		}

		TEST_METHOD(ReadsOriginalXMLBytesFromTheBeginning)
		{
			const std::vector<BYTE> expected(8200, 'x');
			for (const bool partialAsFalse : {false, true})
			{
				TestStream stream(expected);
				stream.partialAsFalse = partialAsFalse;
				response = &stream;
				std::vector<BYTE> actual{1, 2, 3};
				Assert::AreEqual(S_OK, mapi::GetAutoDiscoverXML(L"user@example.invalid", actual));
				Assert::IsTrue(expected == actual);
				Assert::AreEqual(std::wstring(L"user@example.invalid"), requestedAddress);
				Assert::AreEqual(1UL, stream.seeks);
				Assert::AreEqual(partialAsFalse ? 3UL : 4UL, stream.reads);
				Assert::AreEqual(1UL, stream.refs);
			}
			Assert::AreEqual(2UL, requests);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(ServiceFailureReleasesReturnedStreamAndClearsOutput)
		{
			TestStream stream({'x'});
			response = &stream;
			responseResult = MAPI_E_NOT_FOUND;
			std::vector<BYTE> actual{1};
			Assert::AreEqual(MAPI_E_NOT_FOUND, mapi::GetAutoDiscoverXML(L"user@example.invalid", actual));
			Assert::IsTrue(actual.empty());
			Assert::AreEqual(1UL, stream.refs);
			Assert::AreEqual(0UL, stream.reads);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(ReportsMissingStreamAndInvalidAddress)
		{
			std::vector<BYTE> actual{1};
			Assert::AreEqual(MAPI_E_CALL_FAILED, mapi::GetAutoDiscoverXML(L"user@example.invalid", actual));
			Assert::IsTrue(actual.empty());
			Assert::AreEqual(MAPI_E_INVALID_PARAMETER, mapi::GetAutoDiscoverXML(L"", actual));
			Assert::AreEqual(1UL, requests);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(StreamFailuresPreserveErrorsAndReleaseOwnership)
		{
			for (const bool failSeek : {false, true})
			{
				TestStream stream({'x'});
				if (failSeek)
					stream.seekResult = STG_E_SEEKERROR;
				else
					stream.readResult = STG_E_READFAULT;
				response = &stream;
				std::vector<BYTE> actual{1};
				Assert::AreEqual(
					failSeek ? STG_E_SEEKERROR : STG_E_READFAULT,
					mapi::GetAutoDiscoverXML(L"user@example.invalid", actual));
				Assert::IsTrue(actual.empty());
				Assert::AreEqual(1UL, stream.refs);
			}
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(DecodesUTF8WithoutChangingOriginalBytes)
		{
			for (const bool bom : {false, true})
			{
				std::vector<BYTE> xml{'<', 'x', '>', 0xce, 0xa9, '<', '/', 'x', '>'};
				if (bom) xml.insert(xml.begin(), {0xef, 0xbb, 0xbf});
				const auto original = xml;
				std::wstring text;
				Assert::AreEqual(S_OK, mapi::AutoDiscoverXMLToText(xml, text));
				Assert::AreEqual(std::wstring(L"<x>\u03a9</x>"), text);
				Assert::IsTrue(xml == original);
			}
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(DecodesUTF8SupplementaryCharactersAndLineBreaks)
		{
			const std::vector<BYTE> xml{'<', 'x', '>', '\r', '\n', 0xf0, 0x9f, 0x98, 0x80, '<', '/', 'x', '>'};
			const auto original = xml;
			std::wstring text;
			Assert::AreEqual(S_OK, mapi::AutoDiscoverXMLToText(xml, text));
			Assert::AreEqual(std::wstring(L"<x>\r\n\U0001f600</x>"), text);
			Assert::IsTrue(xml == original);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(RejectsEmptyAndMalformedTextExplicitly)
		{
			for (const auto& xml :
				 {std::vector<BYTE>{},
				  std::vector<BYTE>{0xff},
				  std::vector<BYTE>{0xff, 0xfe, '<'},
				  std::vector<BYTE>{'<', 0, 0, 0},
				  std::vector<BYTE>{0xff, 0xfe, 0, 0xd8},
				  std::vector<BYTE>{0xc0, 0xaf},
				  std::vector<BYTE>{0xed, 0xa0, 0x80},
				  std::vector<BYTE>{0xf0, 0x9f},
				  std::vector<BYTE>{0xef, 0xbb, 0xbf}})
			{
				std::wstring text{L"stale"};
				Assert::IsTrue(FAILED(mapi::AutoDiscoverXMLToText(xml, text)));
				Assert::IsTrue(text.empty());
			}
			Assert::IsTrue(m_errors.empty());
		}

	private:
		std::unique_ptr<unittest::testMapiMemory> m_memory;
		std::function<void(const std::wstring&)> m_originalError;
		std::vector<std::wstring> m_errors;
	};
} // namespace autodiscovertest
