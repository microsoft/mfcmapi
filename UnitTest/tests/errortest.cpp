#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/utility/error.h>
#include <core/utility/output.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>

namespace errortest
{
	TEST_CLASS(errortest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD_INITIALIZE(captureDiagnostics)
		{
			m_lastError = GetLastError();
			m_debugTag = registry::debugTag;
			m_displayError = error::displayError;
			m_outputToDbgView = output::outputToDbgView;
			m_messages.clear();
			m_output.clear();
			registry::debugTag = 0;
			error::displayError = [&](const std::wstring& message) {
				m_messages.push_back(message);
				SetLastError(ERROR_INVALID_DATA);
			};
			output::outputToDbgView = [&](const std::wstring& message) {
				m_output += message;
				SetLastError(ERROR_INVALID_DATA);
			};
		}

		TEST_METHOD_CLEANUP(restoreDiagnostics)
		{
			error::displayError = m_displayError;
			output::outputToDbgView = m_outputToDbgView;
			registry::debugTag = m_debugTag;
			SetLastError(m_lastError);
		}

		TEST_METHOD(GetPropsResultMacrosAcceptOnlyErrorsReturnedWarning)
		{
			registry::debugTag = static_cast<DWORD>(output::dbgLevel::HRes);
			for (const auto raw : {S_OK, MAPI_W_ERRORS_RETURNED, S_FALSE, E_FAIL})
			{
				auto calls = 0;
				const auto invoke = [&]() {
					++calls;
					return raw;
				};
				const auto expected = raw == MAPI_W_ERRORS_RETURNED ? S_OK : raw;
				for (const auto dialog : {false, true})
				{
					calls = 0;
					m_messages.clear();
					m_output.clear();
					const auto result = dialog ? EC_H_GETPROPS(invoke()) : WC_H_GETPROPS(invoke());
					Assert::AreEqual<HRESULT>(expected, result);
					Assert::AreEqual(1, calls);
					Assert::AreEqual(dialog && expected != S_OK ? size_t{1} : size_t{0}, m_messages.size());
					Assert::AreEqual(expected == S_OK, m_output.empty());
				}
			}
		}

		TEST_METHOD(GetPropsDiscardMacrosSuppressOnlyErrorsReturnedWarning)
		{
			registry::debugTag = static_cast<DWORD>(output::dbgLevel::HRes);
			for (const auto raw : {S_OK, MAPI_W_ERRORS_RETURNED, S_FALSE, E_FAIL})
			{
				auto calls = 0;
				const auto invoke = [&]() {
					++calls;
					return raw;
				};
				const auto report = raw != S_OK && raw != MAPI_W_ERRORS_RETURNED;
				for (const auto dialog : {false, true})
				{
					calls = 0;
					m_messages.clear();
					m_output.clear();
					if (dialog)
					{
						EC_H_GETPROPS_S(invoke());
					}
					else
					{
						WC_H_GETPROPS_S(invoke());
					}

					Assert::AreEqual(1, calls);
					Assert::AreEqual(dialog && report ? size_t{1} : size_t{0}, m_messages.size());
					Assert::AreEqual(!report, m_output.empty());
				}
			}
		}

		TEST_METHOD(BoolFailuresPreserveFailureAndCapturedDiagnostics)
		{
			for (const auto debug : {false, true})
			{
				registry::debugTag = debug ? static_cast<DWORD>(output::dbgLevel::HRes) : 0;
				for (const auto lastError : {DWORD{ERROR_SUCCESS}, DWORD{ERROR_ACCESS_DENIED}})
				{
					for (const auto dialog : {false, true})
					{
						for (const auto discard : {false, true})
						{
							auto calls = 0;
							const auto fail = [&]() -> BOOL {
								++calls;
								SetLastError(lastError);
								return FALSE;
							};
							m_messages.clear();
							m_output.clear();
							if (discard)
							{
								if (dialog)
								{
									EC_B_S(fail());
								}
								else
								{
									WC_B_S(fail());
								}
							}
							else
							{
								const auto result = dialog ? EC_B(fail()) : WC_B(fail());
								Assert::AreEqual<HRESULT>(lastError ? HRESULT_FROM_WIN32(lastError) : E_FAIL, result);
								Assert::IsTrue(FAILED(result));
							}

							Assert::AreEqual(1, calls);
							Assert::AreEqual(dialog ? size_t{1} : size_t{0}, m_messages.size());
							Assert::AreEqual(!debug, m_output.empty());
							const auto message = lastError ? strings::formatmessagesys(lastError)
														   : strings::loadstring(IDS_API_FAILED_NO_ERROR);
							Assert::IsFalse(message.empty());
							const auto expectedCode = lastError ? L"0x80070005" : L"0x80004005";
							if (dialog)
							{
								Assert::IsTrue(m_messages[0].find(message) != std::wstring::npos);
								Assert::IsTrue(m_messages[0].find(expectedCode) != std::wstring::npos);
							}

							if (debug)
							{
								Assert::IsTrue(m_output.find(strings::StripCarriage(message)) != std::wstring::npos);
								Assert::IsTrue(m_output.find(expectedCode) != std::wstring::npos);
							}
						}
					}
				}
			}
		}

		TEST_METHOD(BoolSuccessIgnoresStaleLastError)
		{
			registry::debugTag = static_cast<DWORD>(output::dbgLevel::HRes);
			auto calls = 0;
			const auto succeed = [&]() -> BOOL {
				++calls;
				SetLastError(ERROR_ACCESS_DENIED);
				return TRUE;
			};
			Assert::AreEqual<HRESULT>(S_OK, EC_B(succeed()));
			Assert::AreEqual<HRESULT>(S_OK, WC_B(succeed()));
			EC_B_S(succeed());
			WC_B_S(succeed());
			Assert::AreEqual(4, calls);
			Assert::IsTrue(m_messages.empty());
			Assert::IsTrue(m_output.empty());
		}

		TEST_METHOD(IgnoreMacrosEvaluateEachExpressionOnce)
		{
			registry::debugTag = static_cast<DWORD>(output::dbgLevel::HRes);
			for (const auto raw : {S_OK, MAPI_W_ERRORS_RETURNED, S_FALSE, E_FAIL})
			{
				for (const auto dialog : {false, true})
				{
					auto calls = 0;
					auto lists = 0;
					const auto invoke = [&]() {
						++calls;
						return raw;
					};
					const auto ignored = [&]() {
						++lists;
						return lists == 1 ? std::list<HRESULT>{MAPI_W_ERRORS_RETURNED} : std::list<HRESULT>{};
					};
					m_messages.clear();
					m_output.clear();
					const auto result = dialog ? EC_H_IGNORE(ignored(), invoke()) : WC_H_IGNORE(ignored(), invoke());
					const auto expected = raw == MAPI_W_ERRORS_RETURNED ? S_OK : raw;
					Assert::AreEqual<HRESULT>(expected, result);
					Assert::AreEqual(1, calls);
					Assert::AreEqual(1, lists);
					Assert::AreEqual(dialog && expected != S_OK ? size_t{1} : size_t{0}, m_messages.size());
					Assert::AreEqual(expected == S_OK, m_output.empty());
				}
			}
		}

	private:
		DWORD m_lastError{};
		DWORD m_debugTag{};
		std::function<void(const std::wstring&)> m_displayError;
		std::function<void(const std::wstring&)> m_outputToDbgView;
		std::vector<std::wstring> m_messages;
		std::wstring m_output;
	};
} // namespace errortest
