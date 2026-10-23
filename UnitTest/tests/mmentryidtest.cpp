#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <MrMapi/mmcli.h>
#include <MrMapi/MMEntryID.h>
#include <core/utility/registry.h>

namespace mmentryidtest
{
	template <typename Interface> class PropertyStub : public Interface
	{
	public:
		ULONG refs{1};
		STDMETHODIMP QueryInterface(REFIID, void** object) override
		{
			*object = nullptr;
			return E_NOINTERFACE;
		}
		STDMETHODIMP_(ULONG) AddRef() override { return ++refs; }
		STDMETHODIMP_(ULONG) Release() override { return --refs; }
		STDMETHODIMP GetLastError(HRESULT, ULONG, LPMAPIERROR*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP SaveChanges(ULONG) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP GetProps(LPSPropTagArray, ULONG, ULONG*, LPSPropValue*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP GetPropList(ULONG, LPSPropTagArray*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP OpenProperty(ULONG, LPCIID, ULONG, ULONG, LPUNKNOWN*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP SetProps(ULONG, LPSPropValue, LPSPropProblemArray*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP DeleteProps(LPSPropTagArray, LPSPropProblemArray*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP
		CopyTo(ULONG, LPCIID, LPSPropTagArray, ULONG_PTR, LPMAPIPROGRESS, LPCIID, LPVOID, ULONG, LPSPropProblemArray*)
			override
		{
			return MAPI_E_NO_SUPPORT;
		}
		STDMETHODIMP
		CopyProps(LPSPropTagArray, ULONG_PTR, LPMAPIPROGRESS, LPCIID, LPVOID, ULONG, LPSPropProblemArray*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		STDMETHODIMP GetNamesFromIDs(LPSPropTagArray*, LPGUID, ULONG, ULONG*, LPMAPINAMEID**) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		STDMETHODIMP GetIDsFromNames(ULONG, LPMAPINAMEID*, ULONG, LPSPropTagArray*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
	};

	class Message final : public PropertyStub<IMessage>
	{
	public:
		HRESULT result{S_OK};
		bool missing{};
		bool noValue{};
		ULONG calls{};
		STDMETHODIMP GetProps(LPSPropTagArray tags, ULONG flags, ULONG* count, LPSPropValue* props) override
		{
			++calls;
			Assert::AreEqual(1UL, tags->cValues);
			Assert::AreEqual(static_cast<ULONG>(PR_SUBJECT_W), tags->aulPropTag[0]);
			Assert::AreEqual(static_cast<ULONG>(fMapiUnicode), flags);
			*count = 0;
			*props = nullptr;
			if (FAILED(result) || noValue) return result;
			const auto allocation = MAPIAllocateBuffer(sizeof(SPropValue), reinterpret_cast<LPVOID*>(props));
			if (FAILED(allocation)) return allocation;
			*count = 1;
			**props = {};
			(*props)->ulPropTag = missing ? CHANGE_PROP_TYPE(PR_SUBJECT_W, PT_ERROR) : PR_SUBJECT_W;
			if (missing)
				(*props)->Value.err = MAPI_E_NOT_FOUND;
			else
				(*props)->Value.lpszW = const_cast<LPWSTR>(L"Fixture subject <&>");
			return missing ? MAPI_W_ERRORS_RETURNED : result;
		}
		STDMETHODIMP GetAttachmentTable(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP OpenAttach(ULONG, LPCIID, ULONG, LPATTACH*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP CreateAttach(LPCIID, ULONG, ULONG*, LPATTACH*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP DeleteAttach(ULONG, ULONG_PTR, LPMAPIPROGRESS, ULONG) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP GetRecipientTable(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP ModifyRecipients(ULONG, LPADRLIST) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP SubmitMessage(ULONG) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP SetReadFlag(ULONG) override { return MAPI_E_NO_SUPPORT; }
	};

	class Store final : public PropertyStub<IMsgStore>
	{
	public:
		Message message;
		HRESULT result{S_OK};
		ULONG type{MAPI_MESSAGE};
		bool noObject{};
		ULONG calls{};
		ULONG expectedFlags{};
		STDMETHODIMP
		OpenEntry(ULONG size, LPENTRYID id, LPCIID iid, ULONG flags, ULONG* objectType, LPUNKNOWN* object) override
		{
			++calls;
			Assert::AreEqual(2UL, size);
			const auto bytes = reinterpret_cast<const BYTE*>(id);
			Assert::AreEqual(static_cast<BYTE>(0), bytes[0]);
			Assert::AreEqual(static_cast<BYTE>(0xAB), bytes[1]);
			Assert::IsTrue(iid && *iid == IID_IMessage);
			Assert::AreEqual(expectedFlags, flags, L"Messages must not be opened for modification or deferred errors.");
			*objectType = type;
			*object = noObject ? nullptr : &message;
			if (*object) message.AddRef();
			return result;
		}
		STDMETHODIMP Advise(ULONG, LPENTRYID, ULONG, LPMAPIADVISESINK, ULONG_PTR*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		STDMETHODIMP Unadvise(ULONG_PTR) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP CompareEntryIDs(ULONG, LPENTRYID, ULONG, LPENTRYID, ULONG, ULONG*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		STDMETHODIMP SetReceiveFolder(LPTSTR, ULONG, ULONG, LPENTRYID) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP GetReceiveFolder(LPTSTR, ULONG, ULONG*, LPENTRYID*, LPTSTR*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP GetReceiveFolderTable(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP StoreLogoff(ULONG*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP AbortSubmit(ULONG, LPENTRYID, ULONG) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP GetOutgoingQueue(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP SetLockState(LPMESSAGE, ULONG) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP FinishedMsg(ULONG, ULONG, LPENTRYID) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP NotifyNewMail(LPNOTIFICATION) override { return MAPI_E_NO_SUPPORT; }
	};

	TEST_CLASS(EntryIDTest)
	{
	public:
		static const bool dummy_var = true;
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }
		bool previousNoCache{};
		TEST_METHOD_INITIALIZE(reset)
		{
			previousNoCache = registry::forceMapiNoCache;
			registry::forceMapiNoCache = false;
		}
		TEST_METHOD_CLEANUP(cleanup)
		{
			for (const auto option : cli::g_options)
				option->clear();
			cli::switchUnswitched.clear();
			registry::forceMapiNoCache = previousNoCache;
		}

		cli::OPTIONS Parse(std::deque<std::wstring> args)
		{
			for (const auto option : cli::g_options)
				option->clear();
			cli::OPTIONS options{};
			cli::ParseArgs(options, args, cli::g_options);
			cli::PostParseCheck(options);
			return options;
		}

		std::deque<std::wstring> Request(
			const std::wstring& hex = L"00aB",
			const std::wstring& property = L"PR_SUBJECT_W",
			const std::wstring& store = L"0")
		{
			return {L"-EntryID", hex, L"-Profile", L"Fixture only", L"-Store", store, property};
		}

		TEST_METHOD(ValidRequests)
		{
			for (const auto property : {L"PR_SUBJECT_W", L"pr_subject_w", L"0037001f", L"0x0037001F"})
			{
				const auto options = Parse(Request(L"00AB", property));
				Assert::AreEqual(static_cast<int>(cli::cmdmodeEntryID), options.mode);
				Assert::IsTrue((options.flags & cli::OPT_NEEDMAPILOGON) != 0);
				Assert::IsTrue((options.flags & cli::OPT_NEEDSTORE) == 0, L"No default store fallback.");
			}
			Assert::AreEqual(
				static_cast<int>(cli::cmdmodeEntryID),
				Parse({L"PR_SUBJECT_W", L"-Store", L"0", L"-Pr", L"Fixture only", L"-EntryID", L"00ab"}).mode);
		}

		TEST_METHOD(InvalidHex)
		{
			for (const auto hex : {L"", L"0", L"xyz", L"00gg", L"0x00AB", L"00 AB", L"00:AB"})
				Assert::AreEqual(static_cast<int>(cli::cmdmodeHelp), Parse(Request(hex)).mode);
			Assert::AreEqual(static_cast<int>(cli::cmdmodeHelp), Parse(Request(std::wstring(16386, L'0'))).mode);
			Assert::AreEqual(static_cast<int>(cli::cmdmodeEntryID), Parse(Request(std::wstring(16384, L'0'))).mode);
		}

		TEST_METHOD(InvalidPropertiesAndStoreIndices)
		{
			for (const auto prop :
				 {L"",
				  L"subject",
				  L"PR_UNKNOWN",
				  L"0037",
				  L"0037001Fjunk",
				  L"100000000",
				  L"00000000",
				  L"00370000",
				  L"0037000A",
				  L"0037000D"})
				Assert::AreEqual(static_cast<int>(cli::cmdmodeHelp), Parse(Request(L"00AB", prop)).mode);
			for (const auto store : {L"", L"-1", L"1x", L"0x1", L"2147483648", L"4294967296", L"+1", L" 0"})
				Assert::AreEqual(
					static_cast<int>(cli::cmdmodeHelp), Parse(Request(L"00AB", L"PR_SUBJECT_W", store)).mode);
		}

		TEST_METHOD(RequiredSelectionAndConflictingModes)
		{
			Assert::AreEqual(static_cast<int>(cli::cmdmodeHelp), Parse({L"-EntryID", L"00AB", L"PR_SUBJECT_W"}).mode);
			Assert::AreEqual(
				static_cast<int>(cli::cmdmodeHelp),
				Parse({L"-EntryID", L"00AB", L"-Profile", L"Fixture only", L"PR_SUBJECT_W"}).mode);
			Assert::AreEqual(
				static_cast<int>(cli::cmdmodeHelp),
				Parse({L"-EntryID", L"00AB", L"-Store", L"0", L"PR_SUBJECT_W"}).mode);
			for (const auto option :
				 {L"-Contents",
				  L"-List",
				  L"-Number",
				  L"-Dispids",
				  L"-MSG",
				  L"-MoreProperties",
				  L"-Recurse",
				  L"-Folder",
				  L"-Output",
				  L"-Input"})
			{
				auto args = Request();
				args.push_back(option);
				if (option == std::wstring(L"-Folder") || option == std::wstring(L"-Output") ||
					option == std::wstring(L"-Input"))
					args.push_back(L"unused");
				Assert::AreEqual(static_cast<int>(cli::cmdmodeHelp), Parse(args).mode, option);
			}
		}

		TEST_METHOD(PreserveExistingModesAndPrefixes)
		{
			Assert::AreEqual(static_cast<int>(cli::cmdmodePropTag), Parse({L"PR_SUBJECT_W"}).mode);
			Assert::AreEqual(static_cast<int>(cli::cmdmodeStoreProperties), Parse({L"-Store", L"0"}).mode);
			Assert::AreEqual(static_cast<int>(cli::cmdmodeContents), Parse({L"-Contents"}).mode);
			Assert::IsTrue(cli::GetOption(L"-E", cli::g_options) == &cli::switchError);
			Assert::IsTrue(cli::GetOption(L"-En", cli::g_options) == &cli::switchEncoding);
			Assert::IsTrue(cli::GetOption(L"-Entry", cli::g_options) == &cli::switchEntryID);
		}

		TEST_METHOD(AdditionalOptions)
		{
			auto args = Request();
			args.insert(args.end(), {L"-NoAddins", L"-Verbose", L"-Online", L"-Version", L"16"});
			Assert::AreEqual(static_cast<int>(cli::cmdmodeEntryID), Parse(args).mode);
			args = Request();
			args.push_back(L"PR_MESSAGE_CLASS_W");
			Assert::AreEqual(static_cast<int>(cli::cmdmodeHelp), Parse(args).mode);
			Assert::AreEqual(
				static_cast<int>(cli::cmdmodeHelp),
				Parse({L"-EntryID", L"00AB", L"-Profile", L"", L"-Store", L"0", L"PR_SUBJECT_W"}).mode);
		}

		void CheckLookup(Store & store, HRESULT expected)
		{
			unittest::testMapiMemory memory;
			Parse(Request());
			Assert::AreEqual(expected, DoEntryID(&store));
			Assert::AreEqual(1UL, store.refs, L"The store is borrowed, not released.");
			Assert::AreEqual(1UL, store.message.refs, L"Release the message on every path.");
		}

		TEST_METHOD(ReadSubject)
		{
			Store store;
			CheckLookup(store, S_OK);
			Assert::AreEqual(1UL, store.message.calls);
		}
		TEST_METHOD(MissingProperty)
		{
			Store store;
			store.message.missing = true;
			CheckLookup(store, MAPI_E_NOT_FOUND);
		}
		TEST_METHOD(OnlineRead)
		{
			registry::forceMapiNoCache = true;
			Store store;
			store.expectedFlags = MAPI_NO_CACHE;
			CheckLookup(store, S_OK);
		}
		TEST_METHOD(OpenFailuresAndNonMessages)
		{
			Assert::AreEqual(static_cast<HRESULT>(MAPI_E_NOT_INITIALIZED), DoEntryID(nullptr));
			Store store;
			store.result = MAPI_E_NOT_FOUND;
			CheckLookup(store, MAPI_E_NOT_FOUND);
			store.result = S_OK;
			store.type = MAPI_FOLDER;
			CheckLookup(store, MAPI_E_INVALID_OBJECT);
			store.type = MAPI_MESSAGE;
			store.noObject = true;
			CheckLookup(store, MAPI_E_INVALID_OBJECT);
			Assert::AreEqual(0UL, store.message.calls);
		}
		TEST_METHOD(GetPropsFailures)
		{
			Store store;
			store.message.result = MAPI_E_CALL_FAILED;
			CheckLookup(store, MAPI_E_CALL_FAILED);
			store.message.result = S_OK;
			store.message.noValue = true;
			CheckLookup(store, MAPI_E_CALL_FAILED);
		}
	};
} // namespace mmentryidtest
