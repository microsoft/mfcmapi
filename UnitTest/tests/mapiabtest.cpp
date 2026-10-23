#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/mapiABFunctions.h>
#include <core/mapi/mapiMemory.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <core/utility/error.h>
#include <stdexcept>
#include <thread>

namespace mapiabtest
{
	class TestSession final : public IMAPISession
	{
	public:
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void** object) override
		{
			*object = nullptr;
			return E_NOINTERFACE;
		}
		ULONG STDMETHODCALLTYPE AddRef() override { return 2; }
		ULONG STDMETHODCALLTYPE Release() override { return 1; }
		HRESULT STDMETHODCALLTYPE GetLastError(HRESULT, ULONG, LPMAPIERROR*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetMsgStoresTable(ULONG, LPMAPITABLE* table) override
		{
			++storeTableCalls;
			*table = nullptr;
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE OpenMsgStore(ULONG_PTR, ULONG, LPENTRYID, LPCIID, ULONG, LPMDB*) override
		{
			++storeOpenCalls;
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE OpenAddressBook(ULONG_PTR ui, LPCIID iid, ULONG flags, LPADRBOOK* book) override
		{
			++addressBookCalls;
			Assert::AreEqual(ULONG_PTR{}, ui);
			Assert::IsNull(iid);
			Assert::AreEqual(static_cast<ULONG>(AB_NO_DIALOG), flags);
			*book = addressBook;
			if (*book) (*book)->AddRef();
			return addressBookResult;
		}
		HRESULT STDMETHODCALLTYPE OpenProfileSection(LPMAPIUID, LPCIID, ULONG, LPPROFSECT*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE GetStatusTable(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
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
		HRESULT STDMETHODCALLTYPE MessageOptions(ULONG_PTR, ULONG, LPTSTR, LPMESSAGE) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE QueryDefaultMessageOpt(LPTSTR, ULONG, ULONG*, LPSPropValue*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE EnumAdrTypes(ULONG, ULONG*, LPTSTR**) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE QueryIdentity(ULONG*, LPENTRYID*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE Logoff(ULONG_PTR, ULONG, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SetDefaultStore(ULONG, ULONG, LPENTRYID) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE AdminServices(ULONG, LPSERVICEADMIN*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE
		ShowForm(ULONG_PTR, LPMDB, LPMAPIFOLDER, LPCIID, ULONG_PTR, LPMESSAGE, ULONG, ULONG, ULONG, ULONG, LPSTR)
			override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE PrepareForm(LPCIID, LPMESSAGE, ULONG_PTR*) override { return MAPI_E_NO_SUPPORT; }

		LPADRBOOK addressBook{};
		HRESULT addressBookResult{S_OK};
		ULONG addressBookCalls{};
		ULONG storeTableCalls{};
		ULONG storeOpenCalls{};
	};

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
		HRESULT STDMETHODCALLTYPE
		GetProps(LPSPropTagArray tags, ULONG flags, ULONG* count, LPSPropValue* props) override
		{
			++getPropsCalls;
			Assert::AreEqual(3UL, tags->cValues);
			Assert::AreEqual(static_cast<ULONG>(MAPI_UNICODE), flags);
			*count = static_cast<ULONG>(mailboxProperties.size());
			*props = nullptr;
			if (!mailboxProperties.empty())
			{
				*props =
					mapi::allocate<LPSPropValue>(static_cast<ULONG>(mailboxProperties.size() * sizeof(SPropValue)));
				Assert::IsNotNull(*props);
				std::copy(mailboxProperties.begin(), mailboxProperties.end(), *props);
			}
			return getPropsResult;
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
		HRESULT STDMETHODCALLTYPE
		OpenEntry(ULONG cbEntryID, LPENTRYID addressEntry, LPCIID iid, ULONG flags, ULONG* type, LPUNKNOWN* object)
			override
		{
			++openCalls;
			Assert::AreEqual(1UL, cbEntryID);
			Assert::AreEqual(static_cast<BYTE>(42), *reinterpret_cast<LPBYTE>(addressEntry));
			Assert::IsTrue(IsEqualIID(*iid, IID_IMAPIProp));
			Assert::AreEqual(0UL, flags);
			*type = resolvedObjectType;
			*object = static_cast<IMAPIProp*>(this);
			AddRef();
			return openResult;
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
		HRESULT STDMETHODCALLTYPE
		ResolveName(ULONG_PTR uiParam, ULONG flags, LPTSTR title, LPADRLIST recipients) override
		{
			++resolveCalls;
			Assert::AreEqual(ULONG_PTR{}, uiParam);
			Assert::AreEqual(static_cast<ULONG>(MAPI_UNICODE), flags);
			Assert::IsNull(title);
			Assert::AreEqual(1UL, recipients->cEntries);
			Assert::AreEqual(1UL, recipients->aEntries[0].cValues);
			Assert::AreEqual(static_cast<ULONG>(PR_DISPLAY_NAME_W), recipients->aEntries[0].rgPropVals[0].ulPropTag);
			requestedSMTP = recipients->aEntries[0].rgPropVals[0].Value.lpszW;
			auto& prop = recipients->aEntries[0].rgPropVals[0];
			prop = {};
			prop.ulPropTag = resolvedEntryIDTag;
			prop.Value.bin = {1, &resolvedEntryID};
			return resolveResult;
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
		HRESULT resolveResult{S_OK};
		HRESULT openResult{S_OK};
		HRESULT getPropsResult{S_OK};
		ULONG resolvedEntryIDTag{PR_ENTRYID};
		BYTE resolvedEntryID{42};
		ULONG resolvedObjectType{MAPI_MAILUSER};
		ULONG resolveCalls{};
		ULONG openCalls{};
		ULONG getPropsCalls{};
		std::wstring requestedSMTP;
		std::vector<SPropValue> mailboxProperties;
		ULONG refs() const noexcept { return m_refs; }

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

		TEST_METHOD(SMTPOpenValidatesBeforeCallingSession)
		{
			TestSession session;
			LPMDB mailbox{};
			Assert::AreEqual(
				MAPI_E_INVALID_PARAMETER, mapi::store::OpenMailboxBySMTP(&session, L"user@example.test", nullptr));
			Assert::AreEqual(
				MAPI_E_INVALID_PARAMETER, mapi::store::OpenMailboxBySMTP(nullptr, L"user@example.test", &mailbox));
			Assert::AreEqual(MAPI_E_INVALID_PARAMETER, mapi::store::OpenMailboxBySMTP(&session, L"alias", &mailbox));
			Assert::IsNull(mailbox);
			Assert::AreEqual(0UL, session.addressBookCalls);
			Assert::AreEqual(0UL, session.storeTableCalls);
		}

		TEST_METHOD(SMTPOpenPropagatesAddressBookFailuresWithoutFallback)
		{
			for (const auto result : {MAPI_E_LOGON_FAILED, S_OK})
			{
				TestSession session;
				session.addressBookResult = result;
				LPMDB mailbox{};
				Assert::AreEqual(
					FAILED(result) ? result : MAPI_E_INVALID_PARAMETER,
					mapi::store::OpenMailboxBySMTP(&session, L"user@example.test", &mailbox));
				Assert::IsNull(mailbox);
				Assert::AreEqual(1UL, session.addressBookCalls);
				Assert::AreEqual(0UL, session.storeTableCalls);
				Assert::AreEqual(0UL, session.storeOpenCalls);
			}
		}

		TEST_METHOD(SMTPOpenRejectsWrongRecipientBeforeOpeningAnyStore)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			ConfigureMailbox(book);
			book.mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"other@example.test");
			TestSession session;
			session.addressBook = &book;
			LPMDB mailbox{};
			Assert::AreEqual(
				MAPI_E_NOT_FOUND, mapi::store::OpenMailboxBySMTP(&session, L"user@example.test", &mailbox));
			Assert::IsNull(mailbox);
			Assert::AreEqual(0UL, session.storeTableCalls);
			Assert::AreEqual(0UL, session.storeOpenCalls);
			Assert::AreEqual(1UL, book.refs());
		}

		TEST_METHOD(SMTPOpenRequiresPrimaryExchangeStoreWithoutFallback)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			ConfigureMailbox(book);
			TestSession session;
			session.addressBook = &book;
			LPMDB mailbox{};
			Assert::AreEqual(
				MAPI_E_CALL_FAILED, mapi::store::OpenMailboxBySMTP(&session, L"user@example.test", &mailbox));
			Assert::IsNull(mailbox);
			Assert::AreEqual(1UL, session.storeTableCalls);
			Assert::AreEqual(0UL, session.storeOpenCalls);
			Assert::AreEqual(1UL, book.refs());
		}

		TEST_METHOD(ChecksPrimarySMTPInputBeforeResolution)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			for (const auto& address : {L"", L"alias", L"@example.test", L"user@", L"a@b@c", L"a b@c", L"<a@b>"})
			{
				std::wstring dn = L"stale";
				Assert::AreEqual(MAPI_E_INVALID_PARAMETER, mapi::ab::ResolveExchangeMailbox(&book, address, dn));
				Assert::IsTrue(dn.empty());
			}
			std::wstring dn = L"stale";
			Assert::AreEqual(
				MAPI_E_INVALID_PARAMETER, mapi::ab::ResolveExchangeMailbox(nullptr, L"user@example.test", dn));
			Assert::IsTrue(dn.empty());
			Assert::AreEqual(0UL, book.resolveCalls);
			Assert::IsTrue(mapi::ab::IsMailboxSMTPAddress(L"user@example.test"));
		}

		TEST_METHOD(ResolvesAndVerifiesExactPrimarySMTPWithoutUI)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			ConfigureMailbox(book);
			std::wstring dn = L"stale";
			Assert::AreEqual(S_OK, mapi::ab::ResolveExchangeMailbox(&book, L"User@Example.Test", dn));
			Assert::AreEqual(std::wstring(L"/o=Test/ou=Test/cn=Recipients/cn=User"), dn);
			Assert::AreEqual(std::wstring(L"User@Example.Test"), book.requestedSMTP);
			Assert::AreEqual(1UL, book.resolveCalls);
			Assert::AreEqual(1UL, book.openCalls);
			Assert::AreEqual(1UL, book.getPropsCalls);
			Assert::AreEqual(1UL, book.refs());
		}

		TEST_METHOD(RejectsANRNearMatchesAndNonExchangeRecipients)
		{
			unittest::testMapiMemory memory;
			for (const auto wrongType : {false, true})
			{
				TestAddressBook book;
				ConfigureMailbox(book);
				if (wrongType)
					book.mailboxProperties[0].Value.lpszW = const_cast<LPWSTR>(L"SMTP");
				else
					book.mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"other@example.test");
				std::wstring dn = L"stale";
				Assert::AreEqual(MAPI_E_NOT_FOUND, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::IsTrue(dn.empty());
				Assert::AreEqual(1UL, book.refs());
			}
		}

		TEST_METHOD(RejectsMissingErrorAndEmptyIdentityProperties)
		{
			unittest::testMapiMemory memory;
			for (const auto index : {0U, 1U, 2U})
			{
				TestAddressBook book;
				ConfigureMailbox(book);
				book.mailboxProperties[index].ulPropTag =
					CHANGE_PROP_TYPE(book.mailboxProperties[index].ulPropTag, PT_ERROR);
				book.mailboxProperties[index].Value.err = MAPI_E_NOT_FOUND;
				book.getPropsResult = MAPI_W_ERRORS_RETURNED;
				std::wstring dn = L"stale";
				Assert::AreEqual(MAPI_E_NOT_FOUND, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::IsTrue(dn.empty());
				Assert::AreEqual(1UL, book.refs());
			}
			TestAddressBook book;
			ConfigureMailbox(book);
			book.mailboxProperties[1].Value.lpszW = const_cast<LPWSTR>(L"");
			std::wstring dn;
			Assert::AreEqual(MAPI_E_NOT_FOUND, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
		}

		TEST_METHOD(PreservesResolutionAndOpenFailures)
		{
			unittest::testMapiMemory memory;
			for (const auto failResolve : {false, true})
			{
				TestAddressBook book;
				ConfigureMailbox(book);
				if (failResolve)
					book.resolveResult = MAPI_E_AMBIGUOUS_RECIP;
				else
					book.openResult = MAPI_E_NO_ACCESS;
				std::wstring dn = L"stale";
				Assert::AreEqual(
					failResolve ? MAPI_E_AMBIGUOUS_RECIP : MAPI_E_NO_ACCESS,
					mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::IsTrue(dn.empty());
				Assert::AreEqual(1UL, book.refs());
				Assert::AreEqual(0UL, book.getPropsCalls);
			}
		}

		TEST_METHOD(PreservesGetPropsFailureAndReleasesReturnedBuffer)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			ConfigureMailbox(book);
			book.getPropsResult = MAPI_E_NO_ACCESS;
			std::wstring dn = L"stale";
			Assert::AreEqual(MAPI_E_NO_ACCESS, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::IsTrue(dn.empty());
			Assert::AreEqual(1UL, book.refs());
		}

		TEST_METHOD(RejectsErrorEntryIDAndNonMailUserWithoutReadingIdentity)
		{
			unittest::testMapiMemory memory;
			for (const auto errorEntry : {false, true})
			{
				TestAddressBook book;
				ConfigureMailbox(book);
				if (errorEntry)
					book.resolvedEntryIDTag = CHANGE_PROP_TYPE(PR_ENTRYID, PT_ERROR);
				else
					book.resolvedObjectType = MAPI_DISTLIST;
				std::wstring dn = L"stale";
				Assert::AreEqual(
					errorEntry ? MAPI_E_NOT_FOUND : MAPI_E_INVALID_OBJECT,
					mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::IsTrue(dn.empty());
				Assert::AreEqual(0UL, book.getPropsCalls);
				Assert::AreEqual(1UL, book.refs());
			}
		}

	private:
		void ConfigureMailbox(TestAddressBook & book)
		{
			book.mailboxProperties.resize(3);
			book.mailboxProperties[0].ulPropTag = PR_ADDRTYPE_W;
			book.mailboxProperties[0].Value.lpszW = const_cast<LPWSTR>(L"EX");
			book.mailboxProperties[1].ulPropTag = PR_EMAIL_ADDRESS_W;
			book.mailboxProperties[1].Value.lpszW = const_cast<LPWSTR>(L"/o=Test/ou=Test/cn=Recipients/cn=User");
			book.mailboxProperties[2].ulPropTag = PR_SMTP_ADDRESS_W;
			book.mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"user@example.test");
		}

		HRESULT Display(TestAddressBook & book)
		{
			return mapi::ab::DisplayDetails(&book, nullptr, sizeof(m_entry), &m_entry);
		}

		ENTRYID m_entry{};
		std::vector<std::wstring> m_errors;
		std::function<void(const std::wstring&)> m_originalDisplayError;
	};
} // namespace mapiabtest
