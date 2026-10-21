#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/mapiABFunctions.h>
#include <core/utility/error.h>
#include <stdexcept>
#include <thread>

namespace mapiabtest
{
	class TestAddressBook final : public IAddrBook
	{
	public:
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void** object) override
		{
			if (!object) return E_POINTER;
			*object = nullptr;
			return E_NOINTERFACE;
		}
		ULONG STDMETHODCALLTYPE AddRef() override { return ++m_refs; }
		ULONG STDMETHODCALLTYPE Release() override { return --m_refs; }
		HRESULT STDMETHODCALLTYPE GetLastError(HRESULT, ULONG, LPMAPIERROR*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SaveChanges(ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetProps(LPSPropTagArray, ULONG, ULONG*, LPSPropValue*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE GetPropList(ULONG, LPSPropTagArray*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE OpenProperty(ULONG, LPCIID, ULONG, ULONG, LPUNKNOWN*) override
		{
			return MAPI_E_NO_SUPPORT;
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
		HRESULT STDMETHODCALLTYPE OpenEntry(ULONG, LPENTRYID, LPCIID, ULONG, ULONG*, LPUNKNOWN*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE CompareEntryIDs(ULONG, LPENTRYID, ULONG, LPENTRYID, ULONG, ULONG*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE Advise(ULONG, LPENTRYID, ULONG, LPMAPIADVISESINK, ULONG_PTR*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE Unadvise(ULONG_PTR) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE CreateOneOff(LPTSTR, LPTSTR, LPTSTR, ULONG, ULONG*, LPENTRYID*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE
		NewEntry(ULONG_PTR, ULONG, ULONG, LPENTRYID, ULONG, LPENTRYID, ULONG*, LPENTRYID*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE ResolveName(ULONG_PTR, ULONG, LPTSTR, LPADRLIST) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE Address(ULONG_PTR*, LPADRPARM, LPADRLIST*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE Details(
			ULONG_PTR* uiParam,
			LPFNDISMISS dismiss,
			LPVOID dismissContext,
			ULONG cbEntryID,
			LPENTRYID entryID,
			LPFNBUTTON button,
			LPVOID buttonContext,
			LPTSTR buttonText,
			ULONG flags) override
		{
			++calls;
			owner = *uiParam;
			size = cbEntryID;
			entry = entryID;
			detailsFlags = flags;
			noCallbacks = !dismiss && !dismissContext && !button && !buttonContext && !buttonText;
			if (onDetails) onDetails();
			return result;
		}
		HRESULT STDMETHODCALLTYPE RecipOptions(ULONG_PTR, ULONG, LPADRENTRY) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE QueryDefaultRecipOpt(LPTSTR, ULONG, ULONG*, LPSPropValue*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE GetPAB(ULONG*, LPENTRYID*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SetPAB(ULONG, LPENTRYID) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetDefaultDir(ULONG*, LPENTRYID*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SetDefaultDir(ULONG, LPENTRYID) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetSearchPath(ULONG, LPSRowSet*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SetSearchPath(ULONG, LPSRowSet) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE PrepareRecips(ULONG, LPSPropTagArray, LPADRLIST) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		std::function<void()> onDetails;
		HRESULT result{S_OK};
		ULONG calls{};
		ULONG_PTR owner{};
		ULONG size{};
		LPENTRYID entry{};
		ULONG detailsFlags{};
		bool noCallbacks{};

	private:
		ULONG m_refs{1};
	};

	TEST_CLASS(mapiabtest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD_INITIALIZE(captureErrors)
		{
			m_errors.clear();
			m_originalDisplayError = std::move(error::displayError);
			error::displayError = [this](const std::wstring& text) { m_errors.push_back(text); };
		}

		TEST_METHOD_CLEANUP(restoreErrors) { error::displayError = std::move(m_originalDisplayError); }

		TEST_METHOD(ForwardsModalDetailsArguments)
		{
			TestAddressBook book;
			const auto owner = reinterpret_cast<HWND>(ULONG_PTR{0x1234});
			Assert::AreEqual(S_OK, mapi::ab::DisplayDetails(&book, owner, sizeof(m_entry), &m_entry));
			Assert::AreEqual(1UL, book.calls);
			Assert::AreEqual(ULONG_PTR{0x1234}, book.owner);
			Assert::AreEqual(static_cast<ULONG>(sizeof(m_entry)), book.size);
			Assert::IsTrue(book.entry == &m_entry);
			Assert::AreEqual(static_cast<ULONG>(DIALOG_MODAL), book.detailsFlags);
			Assert::IsTrue(book.noCallbacks);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(RejectsRepeatedReentryWithoutReleasingOuterGuard)
		{
			TestAddressBook book;
			book.onDetails = [&] {
				Assert::AreEqual(MAPI_E_BUSY, Display(book));
				Assert::AreEqual(MAPI_E_BUSY, Display(book));
			};
			Assert::AreEqual(S_OK, Display(book));
			Assert::AreEqual(1UL, book.calls);
			Assert::AreEqual(size_t{2}, m_errors.size());
			book.onDetails = nullptr;
			Assert::AreEqual(S_OK, Display(book));
			Assert::AreEqual(2UL, book.calls);
		}

		TEST_METHOD(RejectsReentryAcrossAddressBooksAndOwners)
		{
			TestAddressBook first;
			TestAddressBook second;
			first.onDetails = [&] {
				Assert::AreEqual(
					MAPI_E_BUSY,
					mapi::ab::DisplayDetails(
						&second, reinterpret_cast<HWND>(ULONG_PTR{0x5678}), sizeof(m_entry), &m_entry));
			};
			Assert::AreEqual(S_OK, Display(first));
			Assert::AreEqual(1UL, first.calls);
			Assert::AreEqual(0UL, second.calls);
			Assert::AreEqual(size_t{1}, m_errors.size());
			Assert::AreEqual(S_OK, Display(second));
			Assert::AreEqual(1UL, second.calls);
		}

		TEST_METHOD(ProviderErrorReleasesGuardAndIsReportedOnce)
		{
			TestAddressBook book;
			book.result = MAPI_E_CALL_FAILED;
			Assert::AreEqual(MAPI_E_CALL_FAILED, Display(book));
			Assert::AreEqual(size_t{1}, m_errors.size());
			book.result = S_OK;
			Assert::AreEqual(S_OK, Display(book));
			Assert::AreEqual(2UL, book.calls);
		}

		TEST_METHOD(CancellationReleasesGuardWithoutErrorDialog)
		{
			TestAddressBook book;
			for (const auto result : {MAPI_E_USER_CANCEL, MAPI_E_CANCEL})
			{
				book.result = result;
				Assert::AreEqual(S_OK, Display(book));
				Assert::IsTrue(m_errors.empty());
				book.result = S_OK;
				Assert::AreEqual(S_OK, Display(book));
			}
			Assert::AreEqual(4UL, book.calls);
		}

		TEST_METHOD(ExceptionReleasesGuard)
		{
			TestAddressBook book;
			book.onDetails = [] { throw std::runtime_error("test provider exception"); };
			Assert::ExpectException<std::runtime_error>([&] { static_cast<void>(Display(book)); });
			book.onDetails = nullptr;
			Assert::AreEqual(S_OK, Display(book));
			Assert::AreEqual(2UL, book.calls);
		}

		TEST_METHOD(NullAddressBookIsReportedWithoutHoldingGuard)
		{
			Assert::AreEqual(
				MAPI_E_INVALID_PARAMETER, mapi::ab::DisplayDetails(nullptr, nullptr, sizeof(m_entry), &m_entry));
			Assert::AreEqual(size_t{1}, m_errors.size());
			TestAddressBook book;
			Assert::AreEqual(S_OK, Display(book));
			Assert::AreEqual(1UL, book.calls);
		}

		TEST_METHOD(IndependentThreadsHaveIndependentGuards)
		{
			TestAddressBook first;
			TestAddressBook second;
			HRESULT secondResult = MAPI_E_CALL_FAILED;
			first.onDetails = [&] {
				std::thread other([&] { secondResult = Display(second); });
				other.join();
			};
			Assert::AreEqual(S_OK, Display(first));
			Assert::AreEqual(S_OK, secondResult);
			Assert::AreEqual(1UL, first.calls);
			Assert::AreEqual(1UL, second.calls);
			Assert::IsTrue(m_errors.empty());
		}

	private:
		HRESULT Display(TestAddressBook & book)
		{
			return mapi::ab::DisplayDetails(&book, nullptr, sizeof(m_entry), &m_entry);
		}

		ENTRYID m_entry{};
		std::vector<std::wstring> m_errors;
		std::function<void(const std::wstring&)> m_originalDisplayError;
	};
} // namespace mapiabtest
