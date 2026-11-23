#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/utility/error.h>
#include <core/utility/output.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>

namespace errortest
{
	static constexpr error::CallSite site{"TestApi()", "error-contract-test.cpp", 123};

	class TestFormContainer final : public IMAPIFormContainer
	{
	public:
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void** object) override
		{
			*object = nullptr;
			return E_NOINTERFACE;
		}
		ULONG STDMETHODCALLTYPE AddRef() override { return ++m_refs; }
		ULONG STDMETHODCALLTYPE Release() override { return --m_refs; }
		HRESULT STDMETHODCALLTYPE GetLastError(HRESULT hRes, ULONG flags, LPMAPIERROR* error) override
		{
			++calls;
			Assert::AreEqual<HRESULT>(MAPI_E_EXTENDED_ERROR, hRes);
			Assert::AreEqual<ULONG>(fMapiUnicode, flags);
			*error = nullptr;
			if (FAILED(result) || !provideError) return result;

			void* buffer = nullptr;
			const auto allocated = WC_MAPI(MAPIAllocateBuffer(sizeof(MAPIERROR), &buffer));
			if (FAILED(allocated)) return allocated;

			*error = static_cast<LPMAPIERROR>(buffer);
			**error = {1, m_error, m_component, static_cast<ULONG>(MAPI_E_NOT_FOUND), 0x123};
			return result;
		}
		HRESULT STDMETHODCALLTYPE InstallForm(ULONG_PTR, ULONG, LPCTSTR) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE RemoveForm(LPCSTR) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE ResolveMessageClass(LPCSTR, ULONG, LPMAPIFORMINFO*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE
		ResolveMultipleMessageClasses(LPSMESSAGECLASSARRAY, ULONG, LPSMAPIFORMINFOARRAY*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE CalcFormPropSet(ULONG, LPMAPIFORMPROPARRAY*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetDisplay(ULONG, LPTSTR*) override { return MAPI_E_NO_SUPPORT; }

		HRESULT result = S_OK;
		bool provideError = false;
		int calls = 0;

	private:
		ULONG m_refs = 1;
		TCHAR m_error[15] = _T("Provider error");
		TCHAR m_component[15] = _T("Test component");
	};

	TEST_CLASS(errortest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD_INITIALIZE(captureErrors)
		{
			m_originalLastError = GetLastError();
			m_originalDebugTag = registry::debugTag;
			registry::debugTag = 0;
			m_originalOutput = std::move(output::outputToDbgView);
			m_output.clear();
			output::outputToDbgView = [this](const std::wstring& text) { m_output += text; };
			m_originalDisplayError = std::move(error::displayError);
			m_dialogs = 0;
			m_messages.clear();
			error::displayError = [this](const std::wstring& text) {
				++m_dialogs;
				m_messages.push_back(text);
			};
		}

		TEST_METHOD_CLEANUP(restoreErrors)
		{
			error::displayError = std::move(m_originalDisplayError);
			output::outputToDbgView = std::move(m_originalOutput);
			registry::debugTag = m_originalDebugTag;
			SetLastError(m_originalLastError);
		}

		TEST_METHOD(ReportingPreservesStatusAndDialogPolicy)
		{
			for (const auto flags : {error::Flags::ShowDialog, error::Flags::None})
			{
				for (const auto raw : {S_OK, S_FALSE, MAPI_W_ERRORS_RETURNED, E_FAIL})
				{
					m_messages.clear();
					const auto hRes = error::ReportResult(raw, {flags}, site);
					Assert::AreEqual<HRESULT>(raw, hRes);
					const size_t expected = flags == error::Flags::ShowDialog && raw != S_OK ? 1 : 0;
					Assert::AreEqual(expected, m_messages.size());
				}
			}
		}

		TEST_METHOD(IgnoringDiagnosticsAndAcceptingResultsAreDifferentPolicies)
		{
			const HRESULT ignored[] = {MAPI_W_ERRORS_RETURNED, MAPI_E_NOT_FOUND};
			for (const auto flags : {error::Flags::None, error::Flags::ShowDialog})
			{
				m_messages.clear();
				for (const auto raw : ignored)
				{
					const auto preserved = error::ReportResult(raw, {flags}, site, ignored);
					const auto accepted =
						error::ReportResult(raw, {flags | error::Flags::AcceptIgnored}, site, ignored);
					Assert::AreEqual<HRESULT>(raw, preserved);
					Assert::AreEqual<HRESULT>(S_OK, accepted);
				}

				Assert::IsTrue(m_messages.empty());
				const auto unmatched =
					error::ReportResult(E_ACCESSDENIED, {flags | error::Flags::AcceptIgnored}, site, ignored);
				Assert::AreEqual<HRESULT>(E_ACCESSDENIED, unmatched);
				Assert::AreEqual(flags == error::Flags::ShowDialog ? size_t{1} : size_t{0}, m_messages.size());
			}
		}

		TEST_METHOD(MapiTracingPrecedesIgnoreMatching)
		{
			registry::debugTag = static_cast<DWORD>(output::dbgLevel::MAPIFunctions);
			const auto flags = error::Flags::TraceMapi | error::Flags::ShowDialog | error::Flags::AcceptIgnored;
			for (const auto useList : {false, true})
			{
				for (const auto discard : {false, true})
				{
					for (const auto initiallyIgnored : {false, true})
					{
						m_messages.clear();
						HRESULT ignoredCode = initiallyIgnored ? E_FAIL : E_ACCESSDENIED;
						std::list<HRESULT> ignoredList{ignoredCode};
						auto traces = 0;
						output::outputToDbgView = [&](const std::wstring& text) {
							if (text.find(L"TestApi()") != std::wstring::npos)
							{
								++traces;
								ignoredCode = initiallyIgnored ? E_ACCESSDENIED : E_FAIL;
								ignoredList.front() = ignoredCode;
							}
						};
						const auto run = [&](const auto& ignored) {
							if (discard)
							{
								error::ReportResult_S(E_FAIL, {flags}, site, ignored);
							}
							else
							{
								Assert::AreEqual<HRESULT>(
									initiallyIgnored ? E_FAIL : S_OK,
									error::ReportResult(E_FAIL, {flags}, site, ignored));
							}
						};
						if (useList)
						{
							run(ignoredList);
						}
						else
						{
							run(std::span<const HRESULT>{&ignoredCode, 1});
						}

						Assert::AreEqual(1, traces);
						Assert::AreEqual(initiallyIgnored ? size_t{1} : size_t{0}, m_messages.size());
						output::outputToDbgView = {};
					}
				}
			}
		}

		TEST_METHOD(MapiTracingRemainsIndependentOfStatusAndIgnoredDiagnostics)
		{
			registry::debugTag = static_cast<DWORD>(output::dbgLevel::MAPIFunctions);
			const HRESULT ignored[] = {E_FAIL};
			for (const auto raw : {S_OK, E_FAIL})
			{
				m_output.clear();
				const auto hRes =
					error::ReportResult(raw, {error::Flags::ShowDialog | error::Flags::TraceMapi}, site, ignored);
				Assert::AreEqual<HRESULT>(raw, hRes);
				Assert::IsTrue(m_output.find(L"TestApi()") != std::wstring::npos);
				Assert::IsTrue(m_output.find(L"error-contract-test.cpp") != std::wstring::npos);
				Assert::IsTrue(m_output.find(L"123") != std::wstring::npos);
				Assert::IsTrue(m_messages.empty());
			}

			m_output.clear();
			const auto hRes = error::ReportResult(E_FAIL, {}, site);
			Assert::AreEqual<HRESULT>(E_FAIL, hRes);
			Assert::IsTrue(m_output.empty());
		}

		TEST_METHOD(CustomMessageAndCallSiteSurviveReporting)
		{
			const auto hRes = error::ReportResult(E_FAIL, {error::Flags::ShowDialog, IDS_USERCANCELLED}, site);
			Assert::AreEqual<HRESULT>(E_FAIL, hRes);
			Assert::AreEqual(size_t{1}, m_messages.size());
			const auto message = strings::loadstring(IDS_USERCANCELLED);
			Assert::IsFalse(message.empty());
			Assert::IsTrue(m_messages[0].find(message) != std::wstring::npos);
			Assert::IsTrue(m_messages[0].find(L"TestApi()") != std::wstring::npos);
			Assert::IsTrue(m_messages[0].find(L"error-contract-test.cpp") != std::wstring::npos);
			Assert::IsTrue(m_messages[0].find(L"123") != std::wstring::npos);
		}

		TEST_METHOD(ReportingFlagsCombineIndependently)
		{
			registry::debugTag =
				static_cast<DWORD>(output::dbgLevel::HRes) | static_cast<DWORD>(output::dbgLevel::MAPIFunctions);
			const HRESULT ignored[] = {MAPI_E_NOT_FOUND};
			const auto trace = strings::formatmessage(IDS_FUNCTION, site.file, site.line, site.expression);
			Assert::IsFalse(trace.empty());
			for (unsigned mask = 0; mask < 16; ++mask)
			{
				const auto flags = static_cast<error::Flags>(mask);
				const auto showDialog = (mask & 1) != 0;
				const auto traceMapi = (mask & 2) != 0;
				const auto systemMessage = (mask & 4) != 0;
				const auto acceptIgnored = (mask & 8) != 0;
				const UINT messageId = systemMessage ? ERROR_ACCESS_DENIED : IDS_USERCANCELLED;
				const auto message =
					systemMessage ? strings::formatmessagesys(messageId) : strings::loadstring(messageId);
				Assert::IsFalse(message.empty());
				for (const auto raw : {S_OK, S_FALSE, E_FAIL, MAPI_E_NOT_FOUND})
				{
					m_messages.clear();
					m_output.clear();
					const auto hRes = error::ReportResult(raw, {flags, messageId}, site, ignored);
					const auto matched = raw == MAPI_E_NOT_FOUND;
					const auto reported = raw != S_OK && !matched;
					Assert::AreEqual<HRESULT>(matched && acceptIgnored ? S_OK : raw, hRes);
					Assert::AreEqual(reported && showDialog ? size_t{1} : size_t{0}, m_messages.size());
					Assert::AreEqual(traceMapi, m_output.find(trace) != std::wstring::npos);
					Assert::AreEqual(reported, m_output.find(strings::StripCarriage(message)) != std::wstring::npos);
					if (!m_messages.empty())
					{
						Assert::IsTrue(m_messages[0].find(message) != std::wstring::npos);
					}
				}
			}
		}

		TEST_METHOD(BoolSuccessIgnoresStaleLastErrorAndEvaluatesOnce)
		{
			for (const auto lastError : {ERROR_SUCCESS, ERROR_ACCESS_DENIED})
			{
				auto calls = 0;
				const auto hRes = error::CheckBool(
					[&]() -> BOOL {
						++calls;
						SetLastError(lastError);
						return TRUE;
					},
					error::Flags::ShowDialog,
					site);
				Assert::AreEqual(1, calls);
				Assert::AreEqual<HRESULT>(S_OK, hRes);
				Assert::IsTrue(m_messages.empty());
				Assert::IsTrue(m_output.empty());
			}

			const auto boolean = error::CheckBool([] { return true; }, error::Flags::ShowDialog, site);
			Assert::AreEqual<HRESULT>(S_OK, boolean);
		}

		TEST_METHOD(BoolFailuresUseCapturedErrorsAndRespectDiagnosticGating)
		{
			for (const auto tag : {DWORD{0}, static_cast<DWORD>(output::dbgLevel::HRes)})
			{
				registry::debugTag = tag;
				for (const auto flags :
					 {error::Flags::None,
					  error::Flags::ShowDialog,
					  error::Flags::SystemMessage,
					  error::Flags::ShowDialog | error::Flags::SystemMessage})
				{
					for (const auto lastError : {ERROR_SUCCESS, ERROR_ACCESS_DENIED})
					{
						m_messages.clear();
						m_output.clear();
						auto calls = 0;
						const auto hRes = error::CheckBool(
							[&]() -> BOOL {
								++calls;
								SetLastError(lastError);
								return FALSE;
							},
							flags,
							site);
						const auto expected = lastError == ERROR_SUCCESS ? E_FAIL : HRESULT_FROM_WIN32(lastError);
						Assert::AreEqual(1, calls);
						Assert::AreEqual<HRESULT>(expected, hRes);
						Assert::IsTrue(FAILED(hRes));
						const auto showDialog = (flags & error::Flags::ShowDialog) == error::Flags::ShowDialog;
						Assert::AreEqual(showDialog ? size_t{1} : size_t{0}, m_messages.size());
						Assert::AreEqual(tag == 0, m_output.empty());
						const auto diagnostic = lastError == ERROR_SUCCESS
													? strings::loadstring(IDS_API_FAILED_NO_ERROR)
													: strings::formatmessagesys(lastError);
						Assert::IsFalse(diagnostic.empty());
						if (!m_messages.empty())
						{
							Assert::IsTrue(m_messages[0].find(diagnostic) != std::wstring::npos);
						}

						if (tag != 0)
						{
							Assert::IsTrue(m_output.find(strings::StripCarriage(diagnostic)) != std::wstring::npos);
						}
					}
				}
			}
		}

		TEST_METHOD(ReportingCannotReplaceTheCapturedLastError)
		{
			error::displayError = [this](const std::wstring& text) {
				m_messages.push_back(text);
				SetLastError(ERROR_INVALID_DATA);
			};
			for (const auto lastError : {ERROR_SUCCESS, ERROR_ACCESS_DENIED})
			{
				const auto hRes = error::CheckBool(
					[&]() -> BOOL {
						SetLastError(lastError);
						return FALSE;
					},
					error::Flags::ShowDialog,
					site);
				const auto expected = lastError == ERROR_SUCCESS ? E_FAIL : HRESULT_FROM_WIN32(lastError);
				Assert::AreEqual<HRESULT>(expected, hRes);
				Assert::AreEqual<DWORD>(ERROR_INVALID_DATA, GetLastError());
			}
		}

		TEST_METHOD(Win32CodeConversionUsesTheReturnedCodeNotLastError)
		{
			for (const auto flags :
				 {error::Flags::None, error::Flags::ShowDialog, error::Flags::ShowDialog | error::Flags::SystemMessage})
			{
				for (const auto code : {ERROR_SUCCESS, ERROR_ACCESS_DENIED})
				{
					m_messages.clear();
					SetLastError(ERROR_INVALID_DATA);
					const auto hRes = error::CheckWin32(code, flags, site);
					Assert::AreEqual<HRESULT>(HRESULT_FROM_WIN32(code), hRes);
					const auto showDialog = (flags & error::Flags::ShowDialog) == error::Flags::ShowDialog;
					Assert::AreEqual(code != ERROR_SUCCESS && showDialog ? size_t{1} : size_t{0}, m_messages.size());
					if (code != ERROR_SUCCESS && (flags & error::Flags::SystemMessage) == error::Flags::SystemMessage)
					{
						const auto message = strings::formatmessagesys(code);
						Assert::IsFalse(message.empty());
						Assert::IsTrue(m_messages[0].find(message) != std::wstring::npos);
					}
				}
			}
		}

		TEST_METHOD(LegacyWin32ErrorsPreserveCapturedFailuresAndDiagnosticPolicy)
		{
			error::displayError = [this](const std::wstring& text) {
				m_messages.push_back(text);
				SetLastError(ERROR_INVALID_DATA);
			};
			output::outputToDbgView = [this](const std::wstring& text) {
				m_output += text;
				SetLastError(ERROR_INVALID_DATA);
			};
			for (const auto tag : {DWORD{0}, static_cast<DWORD>(output::dbgLevel::HRes)})
			{
				registry::debugTag = tag;
				for (const auto dialog : {false, true})
				{
					for (const auto lastError : {ERROR_SUCCESS, ERROR_ACCESS_DENIED})
					{
						m_messages.clear();
						m_output.clear();
						SetLastError(lastError);
						const auto hRes = error::CheckWin32Error(dialog, site.file, site.line, site.expression);
						Assert::AreEqual<HRESULT>(
							lastError == ERROR_SUCCESS ? E_FAIL : HRESULT_FROM_WIN32(lastError), hRes);
						Assert::AreEqual(dialog ? size_t{1} : size_t{0}, m_messages.size());
						Assert::AreEqual(tag == 0, m_output.empty());
						const auto message = lastError == ERROR_SUCCESS ? strings::loadstring(IDS_API_FAILED_NO_ERROR)
																		: strings::formatmessagesys(lastError);
						if (dialog)
						{
							Assert::IsTrue(m_messages[0].find(message) != std::wstring::npos);
							Assert::IsTrue(m_messages[0].find(L"TestApi()") != std::wstring::npos);
						}

						if (tag != 0)
						{
							Assert::IsTrue(m_output.find(strings::StripCarriage(message)) != std::wstring::npos);
						}
					}
				}
			}
		}

		TEST_METHOD(ErrDialogFormatsArgumentsAndWorksWithoutADialogCallback)
		{
			for (const auto tag : {DWORD{0}, static_cast<DWORD>(output::dbgLevel::HRes)})
			{
				registry::debugTag = tag;
				for (const auto dialog : {true, false})
				{
					m_messages.clear();
					m_output.clear();
					if (!dialog)
					{
						error::displayError = {};
					}

					error::ErrDialog(site.file, site.line, IDS_EDPROBLEMARRAY, L"First problem\r\nSecond problem");
					Assert::AreEqual(dialog ? size_t{1} : size_t{0}, m_messages.size());
					Assert::AreEqual(tag == 0, m_output.empty());
					const auto& message = dialog ? m_messages[0] : m_output;
					if (dialog || tag != 0)
					{
						Assert::IsTrue(
							message.find(L"Problem array:\r\nFirst problem\r\nSecond problem") != std::wstring::npos);
						Assert::IsTrue(message.find(L"error-contract-test.cpp") != std::wstring::npos);
						Assert::IsTrue(message.find(L"123") != std::wstring::npos);
					}

					error::displayError = [this](const std::wstring& text) { m_messages.push_back(text); };
				}
			}
		}

		TEST_METHOD(MapiErrorFormattingRespectsAnsiAndUnicodePayloads)
		{
			char ansiError[] = "Provider error";
			char ansiComponent[] = "Test component";
			wchar_t unicodeError[] = L"Provider \u00E9rror";
			wchar_t unicodeComponent[] = L"Test component";
			const MAPIERROR ansi{
				2,
				reinterpret_cast<LPTSTR>(ansiError),
				reinterpret_cast<LPTSTR>(ansiComponent),
				static_cast<ULONG>(MAPI_E_NOT_FOUND),
				0x123};
			const MAPIERROR unicode{
				2,
				reinterpret_cast<LPTSTR>(unicodeError),
				reinterpret_cast<LPTSTR>(unicodeComponent),
				static_cast<ULONG>(MAPI_E_NOT_FOUND),
				0x123};
			const std::wstring suffix = L"\r\nlpszComponent = Test component\r\nulLowLevelError = 0x8004010F = "
										L"MAPI_E_NOT_FOUND\r\nulContext = 0x123";
			Assert::AreEqual(
				L"ulVersion = 0x2\r\nlpszError = Provider error" + suffix, error::MAPIErrToString(0, ansi));
			Assert::AreEqual(
				L"ulVersion = 0x2\r\nlpszError = Provider \u00E9rror" + suffix,
				error::MAPIErrToString(MAPI_UNICODE, unicode));
		}

		TEST_METHOD(ProblemArraysIncludeEntryDetailsAndAllowEmptyArrays)
		{
			SPropProblemArray properties{0, {{0x12, PR_SUBJECT_A, MAPI_E_NOT_FOUND}}};
			STnefProblemArray tnef{0, {{0x33, 0x44, PR_SUBJECT_A, MAPI_E_NOT_FOUND}}};
			Assert::IsTrue(error::ProblemArrayToString(properties).empty());
			Assert::IsTrue(error::TnefProblemArrayToString(tnef).empty());
			properties.cProblem = tnef.cProblem = 1;
			const auto propertyMessage = error::ProblemArrayToString(properties);
			const auto tnefMessage = error::TnefProblemArrayToString(tnef);
			Assert::IsTrue(propertyMessage.starts_with(L"Index: 0x00000012"));
			Assert::IsTrue(tnefMessage.starts_with(L"Component: 0x00000033\r\nAttribute: 0x00000044"));
			for (const auto& message : {propertyMessage, tnefMessage})
			{
				Assert::IsTrue(message.find(L"PR_SUBJECT_A") != std::wstring::npos);
				Assert::IsTrue(message.find(L"0x8004010F = MAPI_E_NOT_FOUND") != std::wstring::npos);
			}
		}

		TEST_METHOD(ExtendedErrorsQueryOnlyWhenNeededAndHandleMissingPayloads)
		{
			unittest::testMapiMemory memory;
			TestFormContainer container;
			LPMAPIFORMCONTAINER form = &container;
			LPMAPIFORMMGR manager = nullptr;
			for (const auto hRes : {S_OK, S_FALSE, E_FAIL})
			{
				m_messages.clear();
				error::CheckExtendedError(hRes, form);
				error::CheckExtendedError(hRes, manager);
				Assert::AreEqual(0, container.calls);
				Assert::AreEqual(hRes == S_OK ? size_t{0} : size_t{2}, m_messages.size());
				if (hRes != S_OK)
				{
					Assert::IsTrue(m_messages[0].find(error::ErrorNameFromErrorCode(hRes)) != std::wstring::npos);
				}
			}

			for (const auto result : {S_OK, E_ACCESSDENIED})
			{
				for (const auto provideError : {false, true})
				{
					m_messages.clear();
					container.calls = 0;
					container.result = result;
					container.provideError = provideError;
					error::CheckExtendedError(MAPI_E_EXTENDED_ERROR, form);
					Assert::AreEqual(1, container.calls);
					Assert::AreEqual(SUCCEEDED(result) && provideError ? size_t{1} : size_t{0}, m_messages.size());
					if (!m_messages.empty())
					{
						Assert::IsTrue(m_messages[0].find(L"Provider error") != std::wstring::npos);
						Assert::IsTrue(m_messages[0].find(L"Test component") != std::wstring::npos);
					}
				}
			}
		}

		TEST_METHOD(HResultMacrosKeepResultDiscardAndNativeContracts)
		{
			const std::list<HRESULT> ignored{E_FAIL};
			static_assert(std::is_same_v<decltype(EC_H(S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(WC_H(S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(EC_MAPI(S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(WC_MAPI(S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(EC_B(TRUE)), HRESULT>);
			static_assert(std::is_same_v<decltype(WC_B(TRUE)), HRESULT>);
			static_assert(std::is_same_v<decltype(EC_W32(ERROR_SUCCESS)), HRESULT>);
			static_assert(std::is_same_v<decltype(WC_W32(ERROR_SUCCESS)), HRESULT>);
			static_assert(std::is_same_v<decltype(EC_H_MSG(0, S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(WC_H_MSG(0, S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(EC_H_IGNORE_RET_MAPI(ignored, S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(WC_H_IGNORE_RET(ignored, S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(EC_H_IGNORE(ignored, S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(WC_H_IGNORE(ignored, S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(EC_H_GETPROPS(S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(WC_H_GETPROPS(S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(EC_H_CANCEL(S_OK)), HRESULT>);
			static_assert(std::is_same_v<decltype(WC_D(int, 17)), int>);
			static_assert(std::is_void_v<decltype(EC_H_S(S_OK))>);
			static_assert(std::is_void_v<decltype(WC_H_S(S_OK))>);
			static_assert(std::is_void_v<decltype(EC_MAPI_S(S_OK))>);
			static_assert(std::is_void_v<decltype(WC_MAPI_S(S_OK))>);
			static_assert(std::is_void_v<decltype(WC_H_MSG_S(0, S_OK))>);
			static_assert(std::is_void_v<decltype(EC_W32_S(ERROR_SUCCESS))>);
			static_assert(std::is_void_v<decltype(WC_W32_S(ERROR_SUCCESS))>);
			static_assert(std::is_void_v<decltype(EC_B_S(TRUE))>);
			static_assert(std::is_void_v<decltype(WC_B_S(TRUE))>);
			static_assert(std::is_void_v<decltype(EC_H_IGNORE_S(ignored, S_OK))>);
			static_assert(std::is_void_v<decltype(WC_H_IGNORE_S(ignored, S_OK))>);
			static_assert(std::is_void_v<decltype(EC_H_GETPROPS_S(S_OK))>);
			static_assert(std::is_void_v<decltype(WC_H_GETPROPS_S(S_OK))>);
			static_assert(std::is_void_v<decltype(EC_H_CANCEL_S(S_OK))>);
			static_assert(std::is_same_v<decltype(EC_D_DIALOG(IDOK)), INT_PTR>);
			Assert::AreEqual<HRESULT>(S_OK, WC_H(S_OK));
		}

		TEST_METHOD(HResultMacroFamiliesPreserveStatusReportingAndSingleEvaluation)
		{
			registry::debugTag = static_cast<DWORD>(output::dbgLevel::MAPIFunctions);
			for (const auto raw : {S_OK, S_FALSE, MAPI_W_ERRORS_RETURNED, E_FAIL})
			{
				auto calls = 0;
				const auto invoke = [&]() {
					++calls;
					return raw;
				};
				struct ResultCase
				{
					std::function<HRESULT()> run;
					bool dialog;
					bool trace;
				};
				const ResultCase cases[] = {
					{[&] { return EC_H(invoke()); }, true, false},
					{[&] { return WC_H(invoke()); }, false, false},
					{[&] { return EC_MAPI(invoke()); }, true, true},
					{[&] { return WC_MAPI(invoke()); }, false, true},
					{[&] { return EC_H_MSG(IDS_USERCANCELLED, invoke()); }, true, false},
					{[&] { return WC_H_MSG(IDS_USERCANCELLED, invoke()); }, false, false}};
				for (const auto& entry : cases)
				{
					calls = 0;
					m_messages.clear();
					m_output.clear();
					const auto hRes = entry.run();
					Assert::AreEqual<HRESULT>(raw, hRes);
					Assert::AreEqual(1, calls);
					Assert::AreEqual(entry.dialog && raw != S_OK ? size_t{1} : size_t{0}, m_messages.size());
					Assert::AreEqual(!entry.trace, m_output.empty());
				}

				struct DiscardCase
				{
					std::function<void()> run;
					bool dialog;
					bool trace;
				};
				const DiscardCase discards[] = {
					{[&] { EC_H_S(invoke()); }, true, false},
					{[&] { WC_H_S(invoke()); }, false, false},
					{[&] { EC_MAPI_S(invoke()); }, true, true},
					{[&] { WC_MAPI_S(invoke()); }, false, true},
					{[&] { WC_H_MSG_S(IDS_USERCANCELLED, invoke()); }, false, false}};
				for (const auto& entry : discards)
				{
					calls = 0;
					m_messages.clear();
					m_output.clear();
					entry.run();
					Assert::AreEqual(1, calls);
					Assert::AreEqual(entry.dialog && raw != S_OK ? size_t{1} : size_t{0}, m_messages.size());
					Assert::AreEqual(!entry.trace, m_output.empty());
				}
			}
		}

		TEST_METHOD(Win32MacrosConvertReturnedCodesOnce)
		{
			for (const auto code : {ERROR_SUCCESS, ERROR_ACCESS_DENIED})
			{
				for (auto variant = 0; variant < 4; ++variant)
				{
					auto calls = 0;
					const auto invoke = [&]() -> DWORD {
						++calls;
						SetLastError(ERROR_INVALID_DATA);
						return code;
					};
					m_messages.clear();
					auto hRes = S_OK;
					switch (variant)
					{
					case 0:
						hRes = EC_W32(invoke());
						break;
					case 1:
						hRes = WC_W32(invoke());
						break;
					case 2:
						EC_W32_S(invoke());
						break;
					case 3:
						WC_W32_S(invoke());
						break;
					}

					Assert::AreEqual(1, calls);
					if (variant < 2)
					{
						Assert::AreEqual<HRESULT>(HRESULT_FROM_WIN32(code), hRes);
					}

					Assert::AreEqual(
						code != ERROR_SUCCESS && variant % 2 == 0 ? size_t{1} : size_t{0}, m_messages.size());
				}
			}
		}

		TEST_METHOD(BoolMacrosCaptureFailureBeforeReportingAndKeepDiscardDiagnostics)
		{
			error::displayError = [this](const std::wstring& text) {
				m_messages.push_back(text);
				SetLastError(ERROR_INVALID_DATA);
			};
			output::outputToDbgView = [this](const std::wstring& text) {
				m_output += text;
				SetLastError(ERROR_INVALID_DATA);
			};
			for (const auto tag : {DWORD{0}, static_cast<DWORD>(output::dbgLevel::HRes)})
			{
				registry::debugTag = tag;
				for (const auto success : {false, true})
				{
					for (const auto lastError : {ERROR_SUCCESS, ERROR_ACCESS_DENIED})
					{
						for (auto variant = 0; variant < 4; ++variant)
						{
							auto calls = 0;
							const auto invoke = [&]() -> BOOL {
								++calls;
								SetLastError(lastError);
								return success ? TRUE : FALSE;
							};
							m_messages.clear();
							m_output.clear();
							auto hRes = S_OK;
							auto sourceLine = 0;
							switch (variant)
							{
							case 0:
								sourceLine = __LINE__ + 1;
								hRes = EC_B(invoke());
								break;
							case 1:
								sourceLine = __LINE__ + 1;
								hRes = WC_B(invoke());
								break;
							case 2:
								sourceLine = __LINE__ + 1;
								EC_B_S(invoke());
								break;
							case 3:
								sourceLine = __LINE__ + 1;
								WC_B_S(invoke());
								break;
							}

							Assert::AreEqual(1, calls);
							if (variant < 2)
							{
								const auto expected = success					   ? S_OK
													  : lastError == ERROR_SUCCESS ? E_FAIL
																				   : HRESULT_FROM_WIN32(lastError);
								Assert::AreEqual<HRESULT>(expected, hRes);
								Assert::AreEqual(success, SUCCEEDED(hRes));
							}

							Assert::AreEqual(!success && variant % 2 == 0 ? size_t{1} : size_t{0}, m_messages.size());
							Assert::AreEqual(success || tag == 0, m_output.empty());
							const auto message = lastError == ERROR_SUCCESS
													 ? strings::loadstring(IDS_API_FAILED_NO_ERROR)
													 : strings::formatmessagesys(lastError);
							Assert::IsFalse(message.empty());
							if (!m_messages.empty())
							{
								Assert::IsTrue(m_messages[0].find(message) != std::wstring::npos);
								Assert::IsTrue(m_messages[0].find(L"invoke()") != std::wstring::npos);
								Assert::IsTrue(m_messages[0].find(L"errortest.cpp") != std::wstring::npos);
								Assert::IsTrue(m_messages[0].ends_with(std::to_wstring(sourceLine)));
							}

							if (!success && tag != 0)
							{
								Assert::IsTrue(m_output.find(strings::StripCarriage(message)) != std::wstring::npos);
							}
						}
					}
				}
			}
		}

		TEST_METHOD(IgnoreMacrosKeepPreserveAndAcceptPoliciesAndEvaluateInputsOnce)
		{
			registry::debugTag = static_cast<DWORD>(output::dbgLevel::MAPIFunctions);
			for (const auto raw : {S_OK, MAPI_W_ERRORS_RETURNED, E_FAIL, MAPI_E_NOT_FOUND})
			{
				auto calls = 0;
				auto lists = 0;
				const auto invoke = [&]() {
					++calls;
					return raw;
				};
				const auto ignored = [&] {
					++lists;
					return std::list<HRESULT>{MAPI_W_ERRORS_RETURNED, E_FAIL};
				};
				struct ResultCase
				{
					std::function<HRESULT()> run;
					bool dialog;
					bool trace;
					bool accept;
				};
				const ResultCase cases[] = {
					{[&] { return EC_H_IGNORE_RET_MAPI(ignored(), invoke()); }, true, true, false},
					{[&] { return WC_H_IGNORE_RET(ignored(), invoke()); }, false, false, false},
					{[&] { return EC_H_IGNORE(ignored(), invoke()); }, true, true, true},
					{[&] { return WC_H_IGNORE(ignored(), invoke()); }, false, true, true}};
				const auto matched = raw == MAPI_W_ERRORS_RETURNED || raw == E_FAIL;
				for (const auto& entry : cases)
				{
					calls = lists = 0;
					m_messages.clear();
					m_output.clear();
					const auto hRes = entry.run();
					Assert::AreEqual<HRESULT>(matched && entry.accept ? S_OK : raw, hRes);
					Assert::AreEqual(1, calls);
					Assert::AreEqual(1, lists);
					Assert::AreEqual(
						entry.dialog && raw != S_OK && !matched ? size_t{1} : size_t{0}, m_messages.size());
					Assert::AreEqual(!entry.trace, m_output.empty());
				}

				for (const auto dialog : {false, true})
				{
					calls = lists = 0;
					m_messages.clear();
					m_output.clear();
					if (dialog)
					{
						EC_H_IGNORE_S(ignored(), invoke());
					}
					else
					{
						WC_H_IGNORE_S(ignored(), invoke());
					}

					Assert::AreEqual(1, calls);
					Assert::AreEqual(1, lists);
					Assert::AreEqual(dialog && raw != S_OK && !matched ? size_t{1} : size_t{0}, m_messages.size());
					Assert::IsFalse(m_output.empty());
				}
			}
		}

		TEST_METHOD(CancellationMacrosNormalizeOnlyCancellationAndKeepItsDiagnostics)
		{
			const auto cancelledMessage = strings::loadstring(IDS_USERCANCELLED);
			Assert::IsFalse(cancelledMessage.empty());
			for (const auto tag :
				 {DWORD{0},
				  static_cast<DWORD>(output::dbgLevel::HRes) | static_cast<DWORD>(output::dbgLevel::MAPIFunctions)})
			{
				registry::debugTag = tag;
				for (const auto raw : {S_OK, S_FALSE, E_FAIL, MAPI_E_CANCEL, MAPI_E_USER_CANCEL})
				{
					const auto cancelled = raw == MAPI_E_CANCEL || raw == MAPI_E_USER_CANCEL;
					for (const auto discard : {false, true})
					{
						auto calls = 0;
						const auto invoke = [&] {
							++calls;
							return raw;
						};
						m_messages.clear();
						m_output.clear();
						if (discard)
						{
							EC_H_CANCEL_S(invoke());
						}
						else
						{
							const auto hRes = EC_H_CANCEL(invoke());
							Assert::AreEqual<HRESULT>(cancelled ? S_OK : raw, hRes);
						}

						Assert::AreEqual(1, calls);
						Assert::AreEqual(raw != S_OK && !cancelled ? size_t{1} : size_t{0}, m_messages.size());
						Assert::AreEqual(tag == 0, m_output.empty());
						Assert::AreEqual(tag != 0 && cancelled, m_output.find(cancelledMessage) != std::wstring::npos);
					}
				}
			}
		}

		TEST_METHOD(GetPropsReturnsSuccessOnlyForAcceptedResults)
		{
			for (const auto result : {S_OK, MAPI_W_ERRORS_RETURNED, S_FALSE, MAPI_E_NOT_FOUND})
			{
				m_dialogs = 0;
				auto calls = 0;
				const auto getProps = [&]() -> HRESULT {
					++calls;
					return result;
				};
				const auto expected = result == MAPI_W_ERRORS_RETURNED ? S_OK : result;
				const size_t expectedDialogs = expected == S_OK ? 0 : 1;
				Assert::AreEqual<HRESULT>(expected, EC_H_GETPROPS(getProps()));
				Assert::AreEqual(1, calls);
				Assert::AreEqual(expectedDialogs, m_dialogs);
				Assert::AreEqual<HRESULT>(expected, WC_H_GETPROPS(getProps()));
				Assert::AreEqual(2, calls);
				Assert::AreEqual(expectedDialogs, m_dialogs);
			}
		}

		TEST_METHOD(GetPropsDiscardVariantsOnlyDisplayUnexpectedResults)
		{
			for (const auto result : {S_OK, MAPI_W_ERRORS_RETURNED, S_FALSE, MAPI_E_NOT_FOUND})
			{
				m_dialogs = 0;
				auto calls = 0;
				const auto getProps = [&]() -> HRESULT {
					++calls;
					return result;
				};
				const size_t expectedDialogs = result == S_OK || result == MAPI_W_ERRORS_RETURNED ? 0 : 1;
				EC_H_GETPROPS_S(getProps());
				Assert::AreEqual(1, calls);
				Assert::AreEqual(expectedDialogs, m_dialogs);
				WC_H_GETPROPS_S(getProps());
				Assert::AreEqual(2, calls);
				Assert::AreEqual(expectedDialogs, m_dialogs);
			}
		}

	private:
		std::function<void(const std::wstring&)> m_originalDisplayError;
		std::function<void(const std::wstring&)> m_originalOutput;
		std::vector<std::wstring> m_messages;
		std::wstring m_output;
		DWORD m_originalDebugTag{};
		DWORD m_originalLastError{};
		size_t m_dialogs{};
	};
} // namespace errortest
