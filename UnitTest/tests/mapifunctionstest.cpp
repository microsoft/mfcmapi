#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/mapiABFunctions.h>
#include <core/mapi/mapiBinaryFunctions.h>
#include <core/mapi/mapiMemory.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <core/mapi/interfaces.h>
#include <core/mapi/extraPropTags.h>
#include <core/mapi/columnTags.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <core/utility/error.h>
#include <stdexcept>
#include <thread>

static std::function<
	SCODE(LPCIID, ALLOCATEBUFFER*, ALLOCATEMORE*, FREEBUFFER*, LPVOID, ULONG, ULONG, LPSPropTagArray, LPTABLEDATA*)>
	testCreateTable;

#ifdef _M_IX86
#pragma comment(linker, "/export:CreateTable@36=_TestCreateTable@36")
#else
#pragma comment(linker, "/export:CreateTable=TestCreateTable")
#endif

extern "C" SCODE STDAPICALLTYPE TestCreateTable(
	LPCIID iid,
	ALLOCATEBUFFER* allocate,
	ALLOCATEMORE* more,
	FREEBUFFER* free,
	LPVOID reserved,
	ULONG type,
	ULONG index,
	LPSPropTagArray columns,
	LPTABLEDATA* data)
{
	if (!testCreateTable) return MAPI_E_NO_SUPPORT;

	return testCreateTable(iid, allocate, more, free, reserved, type, index, columns, data);
}

namespace mapifunctionstest
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

		HRESULT STDMETHODCALLTYPE
		OpenMsgStore(ULONG_PTR ui, ULONG size, LPENTRYID entry, LPCIID iid, ULONG flags, LPMDB* store) override
		{
			++storeOpenCalls;
			Assert::AreEqual(ULONG_PTR{}, ui);
			Assert::AreEqual(1UL, size);
			Assert::AreEqual(static_cast<BYTE>(42), *reinterpret_cast<LPBYTE>(entry));
			Assert::IsNull(iid);
			storeOpenFlags = flags;
			*store = openedStore;
			if (*store) (*store)->AddRef();
			return storeOpenResult;
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
		HRESULT STDMETHODCALLTYPE
		OpenEntry(ULONG size, LPENTRYID entry, LPCIID iid, ULONG flags, ULONG* type, LPUNKNOWN* object) override
		{
			if (onOpenEntry) return onOpenEntry(size, entry, iid, flags, type, object);
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
		ULONG storeOpenFlags{};
		LPMDB openedStore{};
		HRESULT storeOpenResult{MAPI_E_NO_SUPPORT};
		std::function<HRESULT(ULONG, LPENTRYID, LPCIID, ULONG, ULONG*, LPUNKNOWN*)> onOpenEntry;
	};

	class TestMailboxStore final : public IMsgStore, public IExchangeManageStoreEx
	{
	public:
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** object) override
		{
			*object = nullptr;
			if (IsEqualIID(iid, IID_IExchangeManageStoreEx))
				*object = static_cast<IExchangeManageStoreEx*>(this);
			else
				return E_NOINTERFACE;
			AddRef();
			return S_OK;
		}

		ULONG STDMETHODCALLTYPE AddRef() override { return ++refs; }
		ULONG STDMETHODCALLTYPE Release() override { return --refs; }
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
		HRESULT STDMETHODCALLTYPE Advise(ULONG, LPENTRYID, ULONG, LPMAPIADVISESINK, ULONG_PTR*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE Unadvise(ULONG_PTR) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE CompareEntryIDs(ULONG, LPENTRYID, ULONG, LPENTRYID, ULONG, ULONG*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE OpenEntry(ULONG, LPENTRYID, LPCIID, ULONG, ULONG*, LPUNKNOWN*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE SetReceiveFolder(LPTSTR, ULONG, ULONG, LPENTRYID) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE GetReceiveFolder(LPTSTR, ULONG, ULONG*, LPENTRYID*, LPTSTR*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE GetReceiveFolderTable(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE StoreLogoff(ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE AbortSubmit(ULONG, LPENTRYID, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetOutgoingQueue(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SetLockState(LPMESSAGE, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE FinishedMsg(ULONG, ULONG, LPENTRYID) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE NotifyNewMail(LPNOTIFICATION) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE CreateStoreEntryID(LPCSTR, LPCSTR, ULONG, ULONG*, LPENTRYID*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE EntryIDFromSourceKey(ULONG, LPBYTE, ULONG, LPBYTE, ULONG*, LPENTRYID*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE GetRights(ULONG, LPENTRYID, ULONG, LPENTRYID, ULONG*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE GetMailboxTable(LPSTR, LPMAPITABLE*, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetPublicFolderTable(LPSTR, LPMAPITABLE*, ULONG) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE
		CreateStoreEntryID2(ULONG count, LPSPropValue props, ULONG flags, ULONG* size, LPENTRYID* entry) override
		{
			++createCalls;
			Assert::AreEqual(expectedFlags, flags);
			Assert::AreEqual(
				static_cast<ULONG>(1 + (expectedServer.empty() ? 0 : 2) + (expectedSMTP.empty() ? 0 : 1)), count);
			ULONG index{};
			Assert::AreEqual(static_cast<ULONG>(PR_PROFILE_MAILBOX), props[index].ulPropTag);
			Assert::AreEqual(
				std::string("/o=Test/ou=Test/cn=Recipients/cn=User"), std::string(props[index++].Value.lpszA));
			if (!expectedServer.empty())
			{
				Assert::AreEqual(static_cast<ULONG>(PR_PROFILE_MDB_DN), props[index].ulPropTag);
				Assert::AreEqual(expectedServer, std::string(props[index++].Value.lpszA));
				Assert::AreEqual(static_cast<ULONG>(PR_FORCE_USE_ENTRYID_SERVER), props[index].ulPropTag);
				Assert::IsTrue(props[index++].Value.b != 0);
			}

			if (!expectedSMTP.empty())
			{
				Assert::AreEqual(static_cast<ULONG>(PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W), props[index].ulPropTag);
				Assert::AreEqual(expectedSMTP, std::wstring(props[index++].Value.lpszW));
			}

			*size = 0;
			*entry = nullptr;
			if (FAILED(createResult)) return createResult;
			const auto hr = MAPIAllocateBuffer(1, reinterpret_cast<LPVOID*>(entry));
			if (FAILED(hr)) return hr;
			*size = 1;
			*reinterpret_cast<LPBYTE>(*entry) = 42;
			return S_OK;
		}

		std::string expectedServer;
		std::wstring expectedSMTP{L"user@example.test"};
		ULONG expectedFlags{};
		HRESULT createResult{S_OK};
		ULONG refs{1};
		ULONG createCalls{};
	};

	class TestAddressBook final : public IAddrBook, public IABContainer, public IMAPITable
	{
	public:
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** object) override
		{
			if (!object) return E_POINTER;
			*object = nullptr;
			if (exposeProperties && iid == IID_IMAPIProp)
			{
				*object = static_cast<IMAPIProp*>(static_cast<IAddrBook*>(this));
				AddRef();
				return S_OK;
			}
			return E_NOINTERFACE;
		}

		ULONG STDMETHODCALLTYPE AddRef() override { return ++m_refs; }
		ULONG STDMETHODCALLTYPE Release() override { return --m_refs; }
		ULONG References() const { return m_refs; }
		HRESULT STDMETHODCALLTYPE GetLastError(HRESULT, ULONG, LPMAPIERROR*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SaveChanges(ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE
		GetProps(LPSPropTagArray tags, ULONG flags, ULONG* count, LPSPropValue* props) override
		{
			if (onGetProps) return onGetProps(tags, flags, count, props);
			++getPropsCalls;
			Assert::AreEqual(5UL, tags->cValues);
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

		HRESULT STDMETHODCALLTYPE
		GetIDsFromNames(ULONG count, LPMAPINAMEID* names, ULONG flags, LPSPropTagArray* tags) override
		{
			if (onGetIDs) return onGetIDs(count, names, flags, tags);
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE
		OpenEntry(ULONG cbEntryID, LPENTRYID addressEntry, LPCIID iid, ULONG flags, ULONG* type, LPUNKNOWN* object)
			override
		{
			if (onOpenEntry) return onOpenEntry(cbEntryID, addressEntry, iid, flags, type, object);
			Assert::AreEqual(1UL, cbEntryID);
			Assert::AreEqual(0UL, flags);
			if (IsEqualIID(*iid, IID_IABContainer))
			{
				++containerOpenCalls;
				const auto index = *reinterpret_cast<LPBYTE>(addressEntry);
				Assert::IsTrue(index < searchContainers.size());
				activeContainer = searchContainers[index];
				*type = MAPI_ABCONT;
				*object = static_cast<IABContainer*>(activeContainer);
				activeContainer->AddRef();
				return containerOpenResult;
			}

			const auto index = static_cast<size_t>(*reinterpret_cast<LPBYTE>(addressEntry) - 42);
			Assert::IsTrue(index < activeContainer->tableRecipients.size());
			auto& recipient = *activeContainer->tableRecipients[index];
			++recipient.openCalls;
			Assert::IsTrue(IsEqualIID(*iid, IID_IMAPIProp));
			*type = recipient.resolvedObjectType;
			*object = static_cast<IAddrBook*>(&recipient);
			recipient.AddRef();
			return recipient.openResult;
		}

		HRESULT STDMETHODCALLTYPE GetContentsTable(ULONG flags, LPMAPITABLE* table) override
		{
			++contentsTableCalls;
			Assert::AreEqual(static_cast<ULONG>(MAPI_UNICODE), flags);
			*table = static_cast<IMAPITable*>(this);
			AddRef();
			tablePosition = 0;
			return contentsTableResult;
		}

		HRESULT STDMETHODCALLTYPE GetHierarchyTable(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SetSearchCriteria(LPSRestriction, LPENTRYLIST, ULONG) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE GetSearchCriteria(ULONG, LPSRestriction*, LPENTRYLIST*, ULONG*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE CreateEntry(ULONG, LPENTRYID, ULONG, LPMAPIPROP*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE CopyEntries(LPENTRYLIST, ULONG_PTR, LPMAPIPROGRESS, ULONG) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE DeleteEntries(LPENTRYLIST, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE Advise(ULONG, LPMAPIADVISESINK, ULONG_PTR*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetStatus(ULONG*, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SetColumns(LPSPropTagArray tags, ULONG flags) override
		{
			Assert::AreEqual(1UL, tags->cValues);
			Assert::AreEqual(static_cast<ULONG>(PR_ENTRYID), tags->aulPropTag[0]);
			Assert::AreEqual(0UL, flags);
			return columnsResult;
		}

		HRESULT STDMETHODCALLTYPE QueryColumns(ULONG, LPSPropTagArray*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetRowCount(ULONG, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SeekRow(BOOKMARK, LONG, LONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SeekRowApprox(ULONG, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE QueryPosition(ULONG*, ULONG*, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE FindRow(LPSRestriction, BOOKMARK, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE Restrict(LPSRestriction restriction, ULONG flags) override
		{
			Assert::AreEqual(0UL, flags);
			Assert::AreEqual(static_cast<ULONG>(RES_PROPERTY), restriction->rt);
			Assert::AreEqual(static_cast<ULONG>(RELOP_EQ), restriction->res.resProperty.relop);
			Assert::AreEqual(static_cast<ULONG>(PR_ANR_W), restriction->res.resProperty.ulPropTag);
			Assert::AreEqual(static_cast<ULONG>(PR_ANR_W), restriction->res.resProperty.lpProp->ulPropTag);
			Assert::AreEqual(requestedSMTP, std::wstring(restriction->res.resProperty.lpProp->Value.lpszW));
			return restrictResult;
		}

		HRESULT STDMETHODCALLTYPE CreateBookmark(BOOKMARK*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE FreeBookmark(BOOKMARK) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SortTable(LPSSortOrderSet, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE QuerySortOrder(LPSSortOrderSet*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE QueryRows(LONG requested, ULONG flags, LPSRowSet* rows) override
		{
			++queryCalls;
			Assert::AreEqual(20L, requested);
			Assert::AreEqual(0UL, flags);
			const auto total =
				accountLookupOnly && requestedSMTP.find(L'@') != std::wstring::npos ? size_t{} : tableRecipients.size();
			const auto count = (std::min)(static_cast<size_t>(requested), total - tablePosition);
			*rows = mapi::allocate<LPSRowSet>(static_cast<ULONG>(CbNewSRowSet(count)));
			Assert::IsNotNull(*rows);
			(*rows)->cRows = static_cast<ULONG>(count);
			for (ULONG index = 0; index < count; ++index)
			{
				auto& row = (*rows)->aRow[index];
				row.cValues = 1;
				row.lpProps = mapi::allocate<LPSPropValue>(sizeof(SPropValue) + 1);
				Assert::IsNotNull(row.lpProps);
				row.lpProps[0].ulPropTag = PR_ENTRYID;
				row.lpProps[0].Value.bin = {1, reinterpret_cast<LPBYTE>(row.lpProps + 1)};
				*row.lpProps[0].Value.bin.lpb = static_cast<BYTE>(42 + tablePosition + index);
			}

			tablePosition += count;
			return queryResult;
		}

		HRESULT STDMETHODCALLTYPE Abort() override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE ExpandRow(ULONG, LPBYTE, ULONG, ULONG, LPSRowSet*, ULONG*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE CollapseRow(ULONG, LPBYTE, ULONG, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE WaitForCompletion(ULONG, ULONG, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetCollapseState(ULONG, ULONG, LPBYTE, ULONG*, LPBYTE*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE SetCollapseState(ULONG, ULONG, LPBYTE, BOOKMARK*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		HRESULT STDMETHODCALLTYPE
		ResolveNames(LPSPropTagArray tags, ULONG flags, LPADRLIST recipients, LPFlagList status) override
		{
			Assert::AreEqual(1UL, tags->cValues);
			Assert::AreEqual(static_cast<ULONG>(PR_ENTRYID), tags->aulPropTag[0]);
			Assert::AreEqual(1UL, status->cFlags);
			Assert::AreEqual(static_cast<ULONG>(MAPI_UNRESOLVED), status->ulFlag[0]);
			const auto hr = ResolveRecipient(0, flags, nullptr, recipients);
			status->ulFlag[0] = accountLookupOnly && requestedSMTP.find(L'@') != std::wstring::npos ? MAPI_UNRESOLVED
																									: resolutionStatus;
			return hr;
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
		HRESULT STDMETHODCALLTYPE ResolveName(ULONG_PTR ui, ULONG flags, LPTSTR title, LPADRLIST recipients) override
		{
			return ResolveRecipient(ui, flags, title, recipients);
		}

		HRESULT ResolveRecipient(ULONG_PTR uiParam, ULONG flags, LPTSTR title, LPADRLIST recipients)
		{
			++resolveCalls;
			Assert::AreEqual(ULONG_PTR{}, uiParam);
			Assert::AreEqual(static_cast<ULONG>(MAPI_UNICODE), flags);
			Assert::IsNull(title);
			Assert::AreEqual(1UL, recipients->cEntries);
			Assert::AreEqual(1UL, recipients->aEntries[0].cValues);
			Assert::AreEqual(static_cast<ULONG>(PR_DISPLAY_NAME_W), recipients->aEntries[0].rgPropVals[0].ulPropTag);
			requestedSMTP = recipients->aEntries[0].rgPropVals[0].Value.lpszW;
			requestedLookups.push_back(requestedSMTP);
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
		HRESULT STDMETHODCALLTYPE GetSearchPath(ULONG flags, LPSRowSet* path) override
		{
			++searchPathCalls;
			Assert::AreEqual(0UL, flags);
			*path = nullptr;
			if (!returnSearchPath) return searchPathResult;
			*path = mapi::allocate<LPSRowSet>(static_cast<ULONG>(CbNewSRowSet(searchContainers.size())));
			Assert::IsNotNull(*path);
			(*path)->cRows = static_cast<ULONG>(searchContainers.size());
			for (ULONG index = 0; index < (*path)->cRows; ++index)
			{
				auto& row = (*path)->aRow[index];
				row.cValues = 1;
				row.lpProps = mapi::allocate<LPSPropValue>(sizeof(SPropValue) + 1);
				Assert::IsNotNull(row.lpProps);
				row.lpProps[0].ulPropTag = searchEntryIDTag;
				row.lpProps[0].Value.bin = {1, reinterpret_cast<LPBYTE>(row.lpProps + 1)};
				*row.lpProps[0].Value.bin.lpb = static_cast<BYTE>(index);
			}

			return searchPathResult;
		}

		HRESULT STDMETHODCALLTYPE SetSearchPath(ULONG, LPSRowSet) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE PrepareRecips(ULONG, LPSPropTagArray, LPADRLIST) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		std::function<void()> onDetails;
		std::function<HRESULT(LPSPropTagArray, ULONG, ULONG*, LPSPropValue*)> onGetProps;
		std::function<HRESULT(ULONG, LPMAPINAMEID*, ULONG, LPSPropTagArray*)> onGetIDs;
		std::function<HRESULT(ULONG, LPENTRYID, LPCIID, ULONG, ULONG*, LPUNKNOWN*)> onOpenEntry;
		HRESULT result{S_OK};
		ULONG calls{};
		ULONG_PTR owner{};
		ULONG size{};
		LPENTRYID entry{};
		ULONG detailsFlags{};
		bool noCallbacks{};
		HRESULT resolveResult{S_OK};
		HRESULT searchPathResult{S_OK};
		HRESULT containerOpenResult{S_OK};
		bool returnSearchPath{true};
		ULONG searchPathCalls{};
		ULONG containerOpenCalls{};
		ULONG searchEntryIDTag{PR_ENTRYID};
		ULONG resolutionStatus{MAPI_RESOLVED};
		std::vector<TestAddressBook*> searchContainers{this};
		TestAddressBook* activeContainer{this};
		std::vector<TestAddressBook*> tableRecipients{this};
		size_t tablePosition{};
		ULONG contentsTableCalls{};
		ULONG queryCalls{};
		HRESULT contentsTableResult{S_OK};
		HRESULT columnsResult{S_OK};
		HRESULT restrictResult{S_OK};
		HRESULT queryResult{S_OK};
		HRESULT openResult{S_OK};
		HRESULT getPropsResult{S_OK};
		ULONG resolvedEntryIDTag{PR_ENTRYID};
		BYTE resolvedEntryID{42};
		ULONG resolvedObjectType{MAPI_MAILUSER};
		ULONG resolveCalls{};
		ULONG openCalls{};
		ULONG getPropsCalls{};
		std::wstring requestedSMTP;
		std::vector<std::wstring> requestedLookups;
		bool accountLookupOnly{};
		std::vector<SPropValue> mailboxProperties;
		ULONG refs() const noexcept { return m_refs; }
		bool exposeProperties{};

	private:
		ULONG m_refs{1};
	};

	namespace
	{
		std::vector<BYTE> OneOff(bool unicode)
		{
			std::vector<BYTE> bytes(24);
			Assert::AreEqual(0, memcpy_s(bytes.data() + 4, bytes.size() - 4, &guid::muidOOP, sizeof(GUID)));
			const DWORD flags = unicode ? MAPI_UNICODE : 0;
			Assert::AreEqual(0, memcpy_s(bytes.data() + 20, bytes.size() - 20, &flags, sizeof(flags)));
			if (unicode)
			{
				const wchar_t values[] = L"Alice\0SMTP\0alice@example.test";
				const auto begin = reinterpret_cast<const BYTE*>(values);
				bytes.insert(bytes.end(), begin, begin + sizeof(values));
			}
			else
			{
				const char values[] = "Alice\0SMTP\0alice@example.test";
				const auto begin = reinterpret_cast<const BYTE*>(values);
				bytes.insert(bytes.end(), begin, begin + sizeof(values));
			}
			return bytes;
		}

		std::vector<BYTE> Wrapped(const std::vector<BYTE>& embedded)
		{
			std::vector<BYTE> bytes(21);
			Assert::AreEqual(0, memcpy_s(bytes.data() + 4, bytes.size() - 4, &guid::WAB_GUID, sizeof(GUID)));
			bytes[20] = 0x85;
			bytes.insert(bytes.end(), embedded.begin(), embedded.end());
			return bytes;
		}

		mapi::binary::BinaryTableRow Decode(std::vector<BYTE>& bytes)
		{
			return mapi::binary::DecodeEntryID({static_cast<ULONG>(bytes.size()), bytes.data()});
		}
	} // namespace

	TEST_CLASS(entryidtest)
	{
	public:
		static const bool dummy_var = true;
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(DecodesAnsiAndUnicodeOneOffMembers)
		{
			for (const auto unicode : {false, true})
			{
				auto bytes = OneOff(unicode);
				const auto member = Decode(bytes);
				Assert::AreEqual(S_OK, member.status);
				Assert::AreEqual(std::wstring(L"Alice"), member.name);
				Assert::AreEqual(std::wstring(L"SMTP"), member.addressType);
				Assert::AreEqual(std::wstring(L"alice@example.test"), member.address);
				Assert::IsTrue(member.entryID == bytes && member.openEntryID == bytes);
				Assert::IsTrue(member.decoded.find(L"alice@example.test") != std::wstring::npos);
			}
		}

		TEST_METHOD(UnwrapsMembersWithoutLosingOriginalEntryID)
		{
			auto original = OneOff(true);
			auto bytes = Wrapped(Wrapped(original));
			const auto member = Decode(bytes);
			Assert::AreEqual(S_OK, member.status);
			Assert::IsTrue(member.entryID == bytes);
			Assert::IsTrue(member.openEntryID == original);
			Assert::AreEqual(std::wstring(L"Alice"), member.name);
		}

		TEST_METHOD(KeepsDirectoryAndUnknownEntriesWithoutInventingContactData)
		{
			std::vector<BYTE> directory(28);
			Assert::AreEqual(0, memcpy_s(directory.data() + 4, directory.size() - 4, &guid::muidEMSAB, sizeof(GUID)));
			const char dn[] = "/o=Example/ou=Users/cn=Directory Member";
			directory.insert(directory.end(), dn, dn + sizeof(dn));
			auto member = Decode(directory);
			Assert::AreEqual(S_OK, member.status);
			Assert::IsTrue(member.name.empty() && member.address.empty());
			Assert::IsTrue(member.decoded.find(L"Directory Member") != std::wstring::npos);
			Assert::IsNull(mapi::binary::OpenEntryID(nullptr, nullptr, member));
			Assert::AreEqual(MAPI_E_NO_SUPPORT, member.status);
			Assert::IsTrue(member.entryID == directory);
			std::vector<BYTE> unknown(20, 0x42);
			member = Decode(unknown);
			Assert::IsTrue(member.entryID == unknown);
			Assert::IsTrue(member.name.empty() && member.address.empty());
		}

		TEST_METHOD(PreservesMalformedAndTruncatedEntries)
		{
			auto bytes = OneOff(true);
			for (size_t size = 0; size < bytes.size(); ++size)
			{
				std::vector<BYTE> truncated(bytes.begin(), bytes.begin() + size);
				const auto member = Decode(truncated);
				Assert::AreEqual(MAPI_E_INVALID_ENTRYID, member.status);
				Assert::IsTrue(member.entryID == truncated);
			}
			auto nested = OneOff(false);
			for (unsigned i = 0; i < 26; ++i)
				nested = Wrapped(nested);
			Assert::AreEqual(MAPI_E_INVALID_ENTRYID, Decode(nested).status);
			Assert::AreEqual(MAPI_E_INVALID_ENTRYID, mapi::binary::DecodeEntryID({20, nullptr}).status);
		}

		TEST_METHOD(ValidatesContactWrapperLengthBeforeLegacyUnwrap)
		{
			std::vector<BYTE> bytes(36);
			Assert::AreEqual(0, memcpy_s(bytes.data() + 4, bytes.size() - 4, &guid::muidContabDLL, sizeof(GUID)));
			ULONG type = CONTAB_USER;
			Assert::AreEqual(0, memcpy_s(bytes.data() + 24, bytes.size() - 24, &type, sizeof(type)));
			const auto embedded = OneOff(false);
			ULONG size = static_cast<ULONG>(embedded.size());
			Assert::AreEqual(0, memcpy_s(bytes.data() + 32, bytes.size() - 32, &size, sizeof(size)));
			bytes.insert(bytes.end(), embedded.begin(), embedded.end());
			Assert::AreEqual(S_OK, Decode(bytes).status);
			size = ULONG_MAX;
			Assert::AreEqual(0, memcpy_s(bytes.data() + 32, bytes.size() - 32, &size, sizeof(size)));
			Assert::AreEqual(MAPI_E_INVALID_ENTRYID, Decode(bytes).status);
		}

		TEST_METHOD(ValidatesPropertyTypesAndPreservesAllRowsAndOwnership)
		{
			std::vector<mapi::binary::BinaryTableRow> members;
			const ULONG tag = PROP_TAG(PT_MV_BINARY, 0x9234);
			SPropValue property{};
			property.ulPropTag = tag;
			Assert::AreEqual(S_OK, mapi::binary::ExtractBinaryTableRows(property, tag, members, true));
			Assert::IsTrue(members.empty());
			property.Value.MVbin.cValues = 1;
			Assert::AreEqual(
				MAPI_E_INVALID_PARAMETER, mapi::binary::ExtractBinaryTableRows(property, tag, members, true));
			auto bytes = OneOff(false);
			SBinary entries[] = {{static_cast<ULONG>(bytes.size()), bytes.data()}, {0, nullptr}, {10, nullptr}};
			property.Value.MVbin = {3, entries};
			Assert::AreEqual(S_OK, mapi::binary::ExtractBinaryTableRows(property, tag, members, true));
			Assert::AreEqual(size_t{3}, members.size());
			Assert::AreEqual(S_OK, members[0].status);
			Assert::AreEqual(MAPI_E_INVALID_ENTRYID, members[1].status);
			Assert::AreEqual(MAPI_E_INVALID_ENTRYID, members[2].status);
			bytes[0] = 42;
			Assert::AreEqual(BYTE{0}, members[0].entryID[0]);
			property.ulPropTag = CHANGE_PROP_TYPE(tag, PT_BINARY);
			Assert::AreEqual(MAPI_E_INVALID_TYPE, mapi::binary::ExtractBinaryTableRows(property, tag, members, true));
			Assert::IsTrue(members.empty());
			property.ulPropTag = PROP_TAG(PT_MV_BINARY, 0x9999);
			Assert::AreEqual(MAPI_E_INVALID_TYPE, mapi::binary::ExtractBinaryTableRows(property, tag, members, true));
			property.ulPropTag = CHANGE_PROP_TYPE(tag, PT_ERROR);
			property.Value.err = MAPI_E_NOT_FOUND;
			Assert::AreEqual(MAPI_E_NOT_FOUND, mapi::binary::ExtractBinaryTableRows(property, tag, members, true));
			property.Value.err = S_OK;
			Assert::AreEqual(MAPI_E_INVALID_TYPE, mapi::binary::ExtractBinaryTableRows(property, tag, members, true));
			property.ulPropTag = tag;
			property.Value.MVbin = {ULONG_MAX, entries};
			Assert::AreEqual(MAPI_E_TOO_COMPLEX, mapi::binary::ExtractBinaryTableRows(property, tag, members, true));
		}

		TEST_METHOD(ReadsBinaryPropertyWithoutNamedPropertyLookup)
		{
			unittest::testMapiMemory memory;
			TestAddressBook message;
			auto bytes = OneOff(false);
			SBinary entry{static_cast<ULONG>(bytes.size()), bytes.data()};
			ULONG tag = 0;
			unsigned mappingCalls = 0;
			message.onGetIDs = [&](ULONG, LPMAPINAMEID*, ULONG, LPSPropTagArray*) -> HRESULT {
				++mappingCalls;
				return E_UNEXPECTED;
			};
			message.onGetProps = [&](LPSPropTagArray tags, ULONG flags, ULONG* count, LPSPropValue* properties) {
				Assert::AreEqual(0UL, flags);
				Assert::AreEqual(1UL, tags->cValues);
				Assert::AreEqual(tag, tags->aulPropTag[0]);
				const auto result = MAPIAllocateBuffer(sizeof(SPropValue), reinterpret_cast<void**>(properties));
				if (FAILED(result)) return result;
				*count = 1;
				**properties = {};
				(*properties)->ulPropTag = tag;
				(*properties)->Value.MVbin = {1, &entry};
				return S_OK;
			};
			std::vector<mapi::binary::BinaryTableRow> rows;
			for (const ULONG id : {0x1234, 0x9234})
			{
				tag = PROP_TAG(PT_MV_BINARY, id);
				for (const auto entryIDs : {false, true})
				{
					Assert::AreEqual(
						S_OK,
						mapi::binary::ReadBinaryTableRows(static_cast<IAddrBook*>(&message), tag, rows, entryIDs));
					Assert::AreEqual(size_t{1}, rows.size());
					Assert::IsTrue(rows[0].entryID == bytes);
					Assert::AreEqual(std::wstring(entryIDs ? L"Alice" : L""), rows[0].name);
				}
			}

			Assert::AreEqual(0U, mappingCalls);
		}

		TEST_METHOD(ReportsBinaryPropertyReadFailuresAndClearsRows)
		{
			unittest::testMapiMemory memory;
			TestAddressBook message;
			const auto tag = PROP_TAG(PT_MV_BINARY, 0x9234);
			std::vector<mapi::binary::BinaryTableRow> rows(1);
			Assert::AreEqual(MAPI_E_INVALID_PARAMETER, mapi::binary::ReadBinaryTableRows(nullptr, tag, rows, false));
			Assert::IsTrue(rows.empty());
			Assert::AreEqual(
				MAPI_E_INVALID_TYPE,
				mapi::binary::ReadBinaryTableRows(
					static_cast<IAddrBook*>(&message), CHANGE_PROP_TYPE(tag, PT_BINARY), rows, false));
			message.onGetProps = [](LPSPropTagArray, ULONG, ULONG*, LPSPropValue*) { return MAPI_E_CALL_FAILED; };
			Assert::AreEqual(
				MAPI_E_CALL_FAILED,
				mapi::binary::ReadBinaryTableRows(static_cast<IAddrBook*>(&message), tag, rows, false));
			message.onGetProps = [](LPSPropTagArray, ULONG, ULONG* count, LPSPropValue* properties) {
				*count = 0;
				*properties = nullptr;
				return S_OK;
			};
			Assert::AreEqual(
				MAPI_E_NOT_FOUND,
				mapi::binary::ReadBinaryTableRows(static_cast<IAddrBook*>(&message), tag, rows, false));
			Assert::IsTrue(rows.empty());
		}

		TEST_METHOD(OpensEmbeddedAddressBookEntryReadOnlyAndPreservesProviderError)
		{
			TestAddressBook book;
			auto original = OneOff(false);
			auto wrapped = Wrapped(original);
			auto member = Decode(wrapped);
			unsigned calls = 0;
			book.onOpenEntry = [&](ULONG size, LPENTRYID entry, LPCIID, ULONG flags, ULONG*, LPUNKNOWN* object) {
				++calls;
				Assert::AreEqual(0UL, flags);
				Assert::AreEqual(static_cast<ULONG>(original.size()), size);
				Assert::IsTrue(memcmp(entry, original.data(), size) == 0);
				*object = nullptr;
				return MAPI_E_LOGON_FAILED;
			};
			Assert::IsNull(mapi::binary::OpenEntryID(&book, nullptr, member));
			Assert::AreEqual(1U, calls);
			Assert::AreEqual(MAPI_E_LOGON_FAILED, member.status);
			Assert::IsTrue(member.entryID == wrapped);
		}

		TEST_METHOD(ReturnsOpenedMemberWithOwnedInterfaceReference)
		{
			TestAddressBook book;
			book.exposeProperties = true;
			auto bytes = OneOff(false);
			auto member = Decode(bytes);
			book.onOpenEntry = [&](ULONG, LPENTRYID, LPCIID, ULONG flags, ULONG* type, LPUNKNOWN* object) {
				Assert::AreEqual(0UL, flags);
				*type = MAPI_MAILUSER;
				*object = static_cast<IAddrBook*>(&book);
				book.AddRef();
				return S_OK;
			};
			const auto object = mapi::binary::OpenEntryID(&book, nullptr, member);
			Assert::IsTrue(object == static_cast<IAddrBook*>(&book));
			Assert::AreEqual(S_OK, member.status);
			Assert::AreEqual(2UL, book.References());
			object->Release();
			Assert::AreEqual(1UL, book.References());
		}

		TEST_METHOD(ReportsSessionFallbackErrorAndNeverRequestsModify)
		{
			TestAddressBook book;
			TestSession session;
			book.onOpenEntry = [](ULONG, LPENTRYID, LPCIID, ULONG flags, ULONG*, LPUNKNOWN* object) {
				Assert::AreEqual(0UL, flags);
				*object = nullptr;
				return MAPI_E_NOT_FOUND;
			};
			auto bytes = OneOff(false);
			auto member = Decode(bytes);
			session.onOpenEntry = [&](ULONG size, LPENTRYID entry, LPCIID, ULONG flags, ULONG*, LPUNKNOWN* object) {
				Assert::AreEqual(0UL, flags);
				Assert::AreEqual(static_cast<ULONG>(bytes.size()), size);
				Assert::IsTrue(memcmp(entry, bytes.data(), size) == 0);
				*object = nullptr;
				return MAPI_E_LOGON_FAILED;
			};
			Assert::IsNull(mapi::binary::OpenEntryID(&book, &session, member));
			Assert::AreEqual(MAPI_E_LOGON_FAILED, member.status);
		}
	};

	class TestMemberTableData final : public ITableData
	{
	public:
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void** object) override
		{
			*object = nullptr;
			return E_NOINTERFACE;
		}
		ULONG STDMETHODCALLTYPE AddRef() override { return ++refs; }
		ULONG STDMETHODCALLTYPE Release() override { return --refs; }
		HRESULT STDMETHODCALLTYPE
		HrGetView(LPSSortOrderSet sort, CALLERRELEASE* callback, ULONG_PTR context, LPMAPITABLE* table) override
		{
			++viewCalls;
			Assert::IsNull(sort);
			Assert::IsNull(callback);
			Assert::AreEqual(ULONG_PTR{}, context);
			*table = nullptr;
			if (FAILED(viewResult)) return viewResult;

			*table = static_cast<IMAPITable*>(&view);
			(*table)->AddRef();
			return S_OK;
		}
		HRESULT STDMETHODCALLTYPE HrModifyRow(LPSRow row) override
		{
			++rowCalls;
			if (FAILED(rowResult)) return rowResult;

			const auto tags = binary ? &columns::sptBinaryCols.tags : &columns::sptEntryIDCols.tags;
			Assert::AreEqual(tags->cValues, row->cValues);
			for (ULONG i = 0; i < row->cValues; ++i)
			{
				Assert::AreEqual(tags->aulPropTag[i], row->lpProps[i].ulPropTag);
			}

			const auto props = row->lpProps;
			const auto rowID = props[columns::eidtagPR_ROWID].Value.l;
			Assert::AreEqual(static_cast<LONG>(members.size() + 1), rowID);
			const auto key = props[columns::eidtagPR_INSTANCE_KEY].Value.bin;
			Assert::AreEqual(static_cast<ULONG>(sizeof(rowID)), key.cb);
			Assert::IsTrue(memcmp(key.lpb, &rowID, sizeof(rowID)) == 0);

			mapi::binary::BinaryTableRow member;
			if (binary)
			{
				Assert::AreEqual(0UL, props[1].Value.bin.cb);
				const auto value = props[4].Value.bin;
				Assert::AreEqual(static_cast<LONG>(value.cb), props[3].Value.l);
				if (value.cb) member.entryID.assign(value.lpb, value.lpb + value.cb);
				members.push_back(std::move(member));
				return S_OK;
			}

			const auto entry = props[columns::eidtagPR_ENTRYID].Value.bin;
			if (entry.cb) member.entryID.assign(entry.lpb, entry.lpb + entry.cb);
			member.name = props[columns::eidtagPR_DISPLAY_NAME].Value.lpszW;
			member.addressType = props[columns::eidtagPR_ADDRTYPE].Value.lpszW;
			member.address = props[columns::eidtagPR_EMAIL_ADDRESS].Value.lpszW;
			member.status = props[columns::eidtagPR_STATUS_CODE].Value.l;
			member.decoded = props[columns::eidtagPR_BODY].Value.lpszW;
			members.push_back(std::move(member));
			return S_OK;
		}
		HRESULT STDMETHODCALLTYPE HrDeleteRow(LPSPropValue) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE HrQueryRow(LPSPropValue, LPSRow*, ULONG*) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE HrEnumRow(ULONG, LPSRow*) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE HrNotify(ULONG, ULONG, LPSPropValue) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE HrInsertRow(ULONG, LPSRow) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE HrModifyRows(ULONG, LPSRowSet) override { return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE HrDeleteRows(ULONG, LPSRowSet, ULONG*) override { return E_NOTIMPL; }

		ULONG refs{};
		bool binary{};
		ULONG rowCalls{}, viewCalls{};
		HRESULT rowResult{S_OK}, viewResult{S_OK};
		TestAddressBook view;
		std::vector<mapi::binary::BinaryTableRow> members;
	};

	TEST_CLASS(binarytabletest)
	{
	public:
		static const bool dummy_var = true;
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }
		TEST_METHOD_INITIALIZE(mockTable)
		{
			testCreateTable = [this](
								  LPCIID iid,
								  ALLOCATEBUFFER* allocate,
								  ALLOCATEMORE* more,
								  FREEBUFFER* free,
								  LPVOID reserved,
								  ULONG type,
								  ULONG index,
								  LPSPropTagArray tags,
								  LPTABLEDATA* data) {
				++m_createCalls;
				Assert::IsTrue(*iid == IID_IMAPITableData);
				Assert::IsTrue(allocate == MAPIAllocateBuffer && more == MAPIAllocateMore && free == MAPIFreeBuffer);
				Assert::IsNull(reserved);
				Assert::AreEqual(static_cast<ULONG>(TBLTYPE_SNAPSHOT), type);
				Assert::AreEqual(PR_ROWID, index);
				const auto expected = m_data.binary ? &columns::sptBinaryCols.tags : &columns::sptEntryIDCols.tags;
				Assert::AreEqual(expected->cValues, tags->cValues);
				for (ULONG i = 0; i < tags->cValues; ++i)
					Assert::AreEqual(expected->aulPropTag[i], tags->aulPropTag[i]);

				*data = nullptr;
				if (FAILED(m_createResult) || m_nullData) return m_createResult;

				*data = &m_data;
				m_data.AddRef();
				return S_OK;
			};
		}
		TEST_METHOD_CLEANUP(restoreTable)
		{
			testCreateTable = {};
			Assert::AreEqual(0UL, m_data.refs);
			Assert::AreEqual(1UL, m_data.view.References());
		}

		TEST_METHOD(PreservesArbitraryBinaryWithoutEntryIDInterpretation)
		{
			unittest::testMapiMemory memory;
			m_data.binary = true;
			std::vector<BYTE> bytes{0, 0xFF, 0x42, 0, 1};
			SBinary values[] = {
				{static_cast<ULONG>(bytes.size()), bytes.data()},
				{0, nullptr},
				{static_cast<ULONG>(bytes.size()), bytes.data()}};
			SPropValue property{};
			property.ulPropTag = PROP_TAG(PT_MV_BINARY, 0x1234);
			property.Value.MVbin = {_countof(values), values};
			std::vector<mapi::binary::BinaryTableRow> members;
			Assert::AreEqual(S_OK, mapi::binary::ExtractBinaryTableRows(property, property.ulPropTag, members, false));
			Assert::AreEqual(size_t{3}, members.size());
			for (const auto& member : members)
			{
				Assert::AreEqual(S_OK, member.status);
				Assert::IsTrue(member.openEntryID.empty() && member.decoded.empty() && member.name.empty());
			}

			LPMAPITABLE table = nullptr;
			Assert::AreEqual(S_OK, mapi::binary::CreateBinaryTable(members, &table, false));
			Assert::IsTrue(m_data.members[0].entryID == bytes && m_data.members[2].entryID == bytes);
			Assert::IsTrue(m_data.members[1].entryID.empty());
			table->Release();
			values[0].lpb = nullptr;
			Assert::AreEqual(
				MAPI_E_INVALID_PARAMETER,
				mapi::binary::ExtractBinaryTableRows(property, property.ulPropTag, members, false));
			Assert::IsTrue(members.empty());
		}

		TEST_METHOD(CreatesEmptyGenericBinaryTable)
		{
			unittest::testMapiMemory memory;
			m_data.binary = true;
			LPMAPITABLE table = nullptr;
			Assert::AreEqual(S_OK, mapi::binary::CreateBinaryTable({}, &table, false));
			Assert::AreEqual(0UL, m_data.rowCalls);
			Assert::AreEqual(1UL, m_data.viewCalls);
			table->Release();
		}

		TEST_METHOD(PassesOwnedValuesForDuplicatesAndMalformedRows)
		{
			unittest::testMapiMemory memory;
			auto bytes = OneOff(true);
			auto member = Decode(bytes);
			std::vector<mapi::binary::BinaryTableRow> members{member, member, {}};
			members[1].name = L"Bob";
			members[1].status = MAPI_E_NOT_FOUND;
			members[2].status = MAPI_E_INVALID_ENTRYID;
			LPMAPITABLE table = nullptr;
			Assert::AreEqual(S_OK, mapi::binary::CreateBinaryTable(members, &table, true));
			Assert::IsNotNull(table);
			members.clear();
			bytes.clear();

			Assert::AreEqual(1UL, m_createCalls);
			Assert::AreEqual(3UL, m_data.rowCalls);
			Assert::AreEqual(1UL, m_data.viewCalls);
			Assert::AreEqual(size_t{3}, m_data.members.size());
			for (size_t i = 0; i < m_data.members.size(); ++i)
			{
				const auto& captured = m_data.members[i];
				if (i < 2)
				{
					Assert::IsTrue(captured.entryID == member.entryID);
					Assert::AreEqual(std::wstring(i ? L"Bob" : L"Alice"), captured.name);
					Assert::AreEqual(member.addressType, captured.addressType);
					Assert::AreEqual(member.address, captured.address);
					Assert::AreEqual(i ? MAPI_E_NOT_FOUND : S_OK, captured.status);
					Assert::AreEqual(member.decoded, captured.decoded);
				}
				else
				{
					Assert::IsTrue(captured.entryID.empty());
					Assert::AreEqual(MAPI_E_INVALID_ENTRYID, captured.status);
				}
			}

			table->Release();
		}

		TEST_METHOD(CreatesEmptyTableWithAllColumns)
		{
			unittest::testMapiMemory memory;
			LPMAPITABLE table = nullptr;
			Assert::AreEqual(S_OK, mapi::binary::CreateBinaryTable({}, &table, true));
			Assert::IsNotNull(table);
			Assert::AreEqual(0UL, m_data.rowCalls);
			Assert::AreEqual(1UL, m_data.viewCalls);
			table->Release();
			Assert::AreEqual(MAPI_E_INVALID_PARAMETER, mapi::binary::CreateBinaryTable({}, nullptr, true));
		}

		TEST_METHOD(PropagatesCreationAndRowAndViewFailures)
		{
			unittest::testMapiMemory memory;
			LPMAPITABLE table = nullptr;
			m_createResult = MAPI_E_NOT_ENOUGH_MEMORY;
			Assert::AreEqual(m_createResult, mapi::binary::CreateBinaryTable({}, &table, true));
			Assert::IsNull(table);
			Assert::AreEqual(0UL, m_data.viewCalls);

			m_createResult = S_OK;
			m_nullData = true;
			Assert::AreEqual(E_UNEXPECTED, mapi::binary::CreateBinaryTable({}, &table, true));
			Assert::IsNull(table);
			m_nullData = false;

			m_data.rowResult = MAPI_E_CALL_FAILED;
			Assert::AreEqual(m_data.rowResult, mapi::binary::CreateBinaryTable({{}, {}}, &table, true));
			Assert::IsNull(table);
			Assert::AreEqual(1UL, m_data.rowCalls);
			Assert::AreEqual(0UL, m_data.viewCalls);
			Assert::AreEqual(0UL, m_data.refs);

			m_data.viewResult = MAPI_E_NO_SUPPORT;
			Assert::AreEqual(m_data.viewResult, mapi::binary::CreateBinaryTable({}, &table, true));
			Assert::IsNull(table);
			Assert::AreEqual(1UL, m_data.viewCalls);
		}

	private:
		TestMemberTableData m_data;
		ULONG m_createCalls{};
		HRESULT m_createResult{S_OK};
		bool m_nullData{};
	};

	TEST_CLASS(mapifunctionstest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD_INITIALIZE(captureErrors)
		{
			m_errors.clear();
			m_originalDisplayError = std::move(error::displayError);
			m_originalServerPrompt = std::move(mapi::store::promptServerName);
			error::displayError = [this](const std::wstring& text) { m_errors.push_back(text); };
		}

		TEST_METHOD_CLEANUP(restoreErrors)
		{
			error::displayError = std::move(m_originalDisplayError);
			mapi::store::promptServerName = std::move(m_originalServerPrompt);
		}

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

		TEST_METHOD(SMTPMailboxLogonOmitsServerAndForceProperties)
		{
			unittest::testMapiMemory memory;
			TestMailboxStore store;
			TestSession session;
			session.openedStore = &store;
			session.storeOpenResult = S_OK;
			const auto mailbox = mapi::store::MailboxLogon(
				&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", L"user@example.test", 0, true);
			Assert::IsTrue(mailbox == static_cast<IMsgStore*>(&store));
			Assert::AreEqual(1UL, store.createCalls);
			Assert::AreEqual(1UL, session.storeOpenCalls);
			Assert::AreEqual(
				static_cast<ULONG>(MDB_NO_DIALOG | MDB_NO_MAIL | MDB_TEMPORARY | MAPI_BEST_ACCESS),
				session.storeOpenFlags & ~static_cast<ULONG>(MDB_ONLINE));
			mailbox->Release();
			Assert::AreEqual(1UL, store.refs);
		}

		TEST_METHOD(ForcedServerMailboxLogonPreservesAdvancedProperties)
		{
			unittest::testMapiMemory memory;
			for (const auto smtp : {L"", L"user@example.test"})
			{
				TestMailboxStore store;
				store.expectedServer = "/cn=Configuration/cn=Servers/cn=Test/cn=Microsoft Private MDB";
				store.expectedSMTP = smtp;
				store.expectedFlags = OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP;
				TestSession session;
				session.openedStore = &store;
				session.storeOpenResult = S_OK;
				const auto mailbox = mapi::store::MailboxLogon(
					&session,
					&store,
					strings::LPCSTRToWstring(store.expectedServer.c_str()),
					L"/o=Test/ou=Test/cn=Recipients/cn=User",
					smtp,
					store.expectedFlags,
					true);
				Assert::IsTrue(mailbox == static_cast<IMsgStore*>(&store));
				Assert::AreEqual(1UL, store.createCalls);
				mailbox->Release();
				Assert::AreEqual(1UL, store.refs);
			}
		}

		TEST_METHOD(AddressBookMailboxOpenUsesSelectedIdentityWithoutServerPrompt)
		{
			unittest::testMapiMemory memory;
			ULONG prompts{};
			mapi::store::promptServerName = [&] {
				++prompts;
				return std::wstring{};
			};
			for (const auto flags :
				 {ULONG{}, static_cast<ULONG>(OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP)})
			{
				TestMailboxStore store;
				store.expectedFlags = flags;
				TestSession session;
				session.openedStore = &store;
				session.storeOpenResult = S_OK;
				const auto mailbox = mapi::store::OpenOtherUsersMailbox(
					&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", L"user@example.test", flags, true);
				Assert::IsTrue(mailbox == static_cast<IMsgStore*>(&store));
				Assert::AreEqual(0UL, prompts);
				Assert::AreEqual(0UL, session.addressBookCalls);
				Assert::AreEqual(1UL, store.createCalls);
				Assert::AreEqual(1UL, session.storeOpenCalls);
				mailbox->Release();
				Assert::AreEqual(1UL, store.refs);
			}
		}

		TEST_METHOD(AddressBookMailboxOpenDoesNotPromptOrFallbackAfterEntryIDFailure)
		{
			unittest::testMapiMemory memory;
			ULONG prompts{};
			mapi::store::promptServerName = [&] {
				++prompts;
				return std::wstring{};
			};
			TestMailboxStore store;
			store.createResult = MAPI_E_NO_ACCESS;
			TestSession session;
			Assert::IsNull(mapi::store::OpenOtherUsersMailbox(
				&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", L"user@example.test", 0, true));
			Assert::AreEqual(0UL, prompts);
			Assert::AreEqual(1UL, store.createCalls);
			Assert::AreEqual(0UL, session.storeOpenCalls);
			Assert::AreEqual(1UL, store.refs);
		}

		TEST_METHOD(DNOnlyMailboxOpenResolvesRoutingSMTPWithoutServerPrompt)
		{
			unittest::testMapiMemory memory;
			ULONG prompts{};
			mapi::store::promptServerName = [&] {
				++prompts;
				return std::wstring{};
			};
			TestAddressBook book;
			ConfigureMailbox(book);
			TestMailboxStore store;
			TestSession session;
			session.addressBook = &book;
			session.openedStore = &store;
			session.storeOpenResult = S_OK;
			const auto mailbox = mapi::store::OpenOtherUsersMailbox(
				&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", {}, 0, true);
			Assert::IsTrue(mailbox == static_cast<IMsgStore*>(&store));
			Assert::AreEqual(0UL, prompts);
			Assert::AreEqual(1UL, session.addressBookCalls);
			Assert::AreEqual(1UL, book.resolveCalls);
			Assert::AreEqual(std::wstring(L"/o=Test/ou=Test/cn=Recipients/cn=User"), book.requestedSMTP);
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(1UL, store.createCalls);
			Assert::AreEqual(1UL, session.storeOpenCalls);
			mailbox->Release();
			Assert::AreEqual(1UL, store.refs);
		}

		TEST_METHOD(DNOnlyMailboxOpenDoesNotPromptOrFallbackAfterProviderFailure)
		{
			unittest::testMapiMemory memory;
			ULONG prompts{};
			mapi::store::promptServerName = [&] {
				++prompts;
				return std::wstring{};
			};
			TestAddressBook book;
			ConfigureMailbox(book);
			TestMailboxStore store;
			store.createResult = MAPI_E_NO_ACCESS;
			TestSession session;
			session.addressBook = &book;
			Assert::IsNull(mapi::store::OpenOtherUsersMailbox(
				&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", {}, 0, true));
			Assert::AreEqual(0UL, prompts);
			Assert::AreEqual(1UL, book.resolveCalls);
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(0UL, session.storeOpenCalls);
			Assert::AreEqual(1UL, store.refs);
		}

		TEST_METHOD(DNMailboxResolutionVerifiesExactRecipientAndReturnsSMTP)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			ConfigureMailbox(book);
			const std::wstring dn = L"/O=TEST/OU=TEST/CN=RECIPIENTS/CN=USER";
			std::wstring smtp = L"stale";
			Assert::AreEqual(S_OK, mapi::ab::ResolveExchangeMailboxSMTP(&book, dn, smtp));
			Assert::AreEqual(std::wstring(L"user@example.test"), smtp);
			Assert::AreEqual(dn, book.requestedSMTP);
			Assert::AreEqual(1UL, book.resolveCalls);
			Assert::AreEqual(1UL, book.openCalls);
			Assert::AreEqual(0UL, book.searchPathCalls);
			Assert::AreEqual(1UL, book.refs());
		}

		TEST_METHOD(DNMailboxResolutionRejectsInvalidInputBeforeLookup)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			std::wstring smtp = L"stale";
			Assert::AreEqual(
				MAPI_E_INVALID_PARAMETER, mapi::ab::ResolveExchangeMailboxSMTP(nullptr, L"/o=Test/cn=User", smtp));
			Assert::IsTrue(smtp.empty());
			for (const auto& dn : {std::wstring{}, std::wstring(L"dn\0suffix", 9)})
			{
				smtp = L"stale";
				Assert::AreEqual(MAPI_E_INVALID_PARAMETER, mapi::ab::ResolveExchangeMailboxSMTP(&book, dn, smtp));
				Assert::IsTrue(smtp.empty());
			}

			Assert::AreEqual(0UL, book.resolveCalls);
		}

		TEST_METHOD(DNMailboxResolutionPreservesProviderFailuresAndCleansUp)
		{
			unittest::testMapiMemory memory;
			for (const auto stage : {0, 1, 2})
			{
				TestAddressBook book;
				ConfigureMailbox(book);
				const auto expected = stage == 0 ? MAPI_E_AMBIGUOUS_RECIP : MAPI_E_NO_ACCESS;
				if (stage == 0) book.resolveResult = expected;
				if (stage == 1) book.openResult = expected;
				if (stage == 2) book.getPropsResult = expected;
				std::wstring smtp = L"stale";
				Assert::AreEqual(
					expected,
					mapi::ab::ResolveExchangeMailboxSMTP(&book, L"/o=Test/ou=Test/cn=Recipients/cn=User", smtp));
				Assert::IsTrue(smtp.empty());
				Assert::AreEqual(1UL, book.refs());
				Assert::AreEqual(stage == 2 ? 1UL : 0UL, book.getPropsCalls);
			}
		}

		TEST_METHOD(DNMailboxResolutionRejectsInvalidResolvedObjects)
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
				std::wstring smtp = L"stale";
				Assert::AreEqual(
					MAPI_E_NOT_FOUND,
					mapi::ab::ResolveExchangeMailboxSMTP(&book, L"/o=Test/ou=Test/cn=Recipients/cn=User", smtp));
				Assert::IsTrue(smtp.empty());
				Assert::AreEqual(0UL, book.getPropsCalls);
				Assert::AreEqual(1UL, book.refs());
			}
		}

		TEST_METHOD(DNMailboxResolutionRejectsMismatchedOrIncompleteIdentity)
		{
			unittest::testMapiMemory memory;
			for (const auto stage : {0, 1, 2, 3, 4, 5, 6, 7})
			{
				TestAddressBook book;
				ConfigureMailbox(book);
				if (stage == 0) book.mailboxProperties[0].Value.lpszW = const_cast<LPWSTR>(L"SMTP");
				if (stage == 1) book.mailboxProperties[1].Value.lpszW = const_cast<LPWSTR>(L"/o=Test/cn=OtherUser");
				if (stage == 2 || stage == 4 || stage == 5)
				{
					const auto index = stage == 2 ? 2U : stage == 4 ? 1U : 0U;
					book.mailboxProperties[index].ulPropTag =
						CHANGE_PROP_TYPE(book.mailboxProperties[index].ulPropTag, PT_ERROR);
					book.mailboxProperties[index].Value.err = MAPI_E_NOT_FOUND;
					book.getPropsResult = MAPI_W_ERRORS_RETURNED;
				}

				if (stage == 3) book.mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"invalid-smtp");
				if (stage == 6) book.mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"");
				if (stage == 7) book.mailboxProperties[1].Value.lpszW = const_cast<LPWSTR>(L"");
				std::wstring smtp = L"stale";
				Assert::AreEqual(
					MAPI_E_NOT_FOUND,
					mapi::ab::ResolveExchangeMailboxSMTP(&book, L"/o=Test/ou=Test/cn=Recipients/cn=User", smtp));
				Assert::IsTrue(smtp.empty());
				Assert::AreEqual(1UL, book.refs());
			}
		}

		TEST_METHOD(DNMailboxOpenStopsOnRoutingLookupFailure)
		{
			unittest::testMapiMemory memory;
			ULONG prompts{};
			mapi::store::promptServerName = [&] {
				++prompts;
				return std::wstring{};
			};
			TestAddressBook book;
			book.resolveResult = MAPI_E_NOT_FOUND;
			TestMailboxStore store;
			TestSession session;
			session.addressBook = &book;
			Assert::IsNull(mapi::store::OpenOtherUsersMailbox(
				&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", {}, 0, true));
			Assert::AreEqual(0UL, prompts);
			Assert::AreEqual(1UL, book.resolveCalls);
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(0UL, store.createCalls);
			Assert::AreEqual(0UL, session.storeOpenCalls);
			Assert::IsFalse(m_errors.empty());
		}

		TEST_METHOD(DNMailboxOpenStopsOnAddressBookFailureOrMissingObject)
		{
			unittest::testMapiMemory memory;
			ULONG prompts{};
			mapi::store::promptServerName = [&] {
				++prompts;
				return std::wstring{};
			};
			for (const auto failOpen : {false, true})
			{
				TestAddressBook book;
				TestMailboxStore store;
				TestSession session;
				if (failOpen)
				{
					session.addressBook = &book;
					session.addressBookResult = MAPI_E_NO_ACCESS;
				}

				Assert::IsNull(mapi::store::OpenOtherUsersMailbox(
					&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", {}, 0, true));
				Assert::AreEqual(0UL, prompts);
				Assert::AreEqual(1UL, session.addressBookCalls);
				Assert::AreEqual(0UL, book.resolveCalls);
				Assert::AreEqual(1UL, book.refs());
				Assert::AreEqual(0UL, store.createCalls);
				Assert::AreEqual(0UL, session.storeOpenCalls);
			}

			Assert::IsFalse(m_errors.empty());
		}

		TEST_METHOD(DNMailboxOpenRejectsMissingRoutingSMTPWithoutFallback)
		{
			unittest::testMapiMemory memory;
			ULONG prompts{};
			mapi::store::promptServerName = [&] {
				++prompts;
				return std::wstring{};
			};
			TestAddressBook book;
			ConfigureMailbox(book);
			book.mailboxProperties[2].ulPropTag = CHANGE_PROP_TYPE(PR_SMTP_ADDRESS_W, PT_ERROR);
			book.mailboxProperties[2].Value.err = MAPI_E_NOT_FOUND;
			book.getPropsResult = MAPI_W_ERRORS_RETURNED;
			TestMailboxStore store;
			TestSession session;
			session.addressBook = &book;
			Assert::IsNull(mapi::store::OpenOtherUsersMailbox(
				&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", {}, 0, true));
			Assert::AreEqual(0UL, prompts);
			Assert::AreEqual(1UL, book.resolveCalls);
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(0UL, store.createCalls);
			Assert::AreEqual(0UL, session.storeOpenCalls);
			Assert::IsFalse(m_errors.empty());
		}

		TEST_METHOD(SMTPMailboxLogonDoesNotOpenOnEntryIDFailure)
		{
			unittest::testMapiMemory memory;
			TestMailboxStore store;
			store.createResult = MAPI_E_NO_ACCESS;
			TestSession session;
			Assert::IsNull(mapi::store::MailboxLogon(
				&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", L"user@example.test", 0, true));
			Assert::AreEqual(1UL, store.createCalls);
			Assert::AreEqual(0UL, session.storeOpenCalls);
			Assert::AreEqual(1UL, store.refs);
		}

		TEST_METHOD(SMTPMailboxLogonCleansUpAfterStoreOpenFailure)
		{
			unittest::testMapiMemory memory;
			TestMailboxStore store;
			TestSession session;
			session.storeOpenResult = MAPI_E_NO_ACCESS;
			Assert::IsNull(mapi::store::MailboxLogon(
				&session, &store, {}, L"/o=Test/ou=Test/cn=Recipients/cn=User", L"user@example.test", 0, true));
			Assert::AreEqual(1UL, store.createCalls);
			Assert::AreEqual(1UL, session.storeOpenCalls);
			Assert::AreEqual(1UL, store.refs);
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

		TEST_METHOD(ChecksSMTPInputBeforeResolution)
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
			Assert::IsTrue(mapi::ab::IsSMTPAddress(L"user@example.test"));
		}

		TEST_METHOD(ResolvesAndVerifiesExactSMTPWithoutUI)
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

		TEST_METHOD(ResolvesAccountIndexedAliasesAndVerifiesUnicodeOrANSIProxyAddresses)
		{
			unittest::testMapiMemory memory;
			for (const auto useTable : {false, true})
			{
				for (const auto unicode : {false, true})
				{
					TestAddressBook book;
					ConfigureMailbox(book);
					book.accountLookupOnly = true;
					if (useTable) book.resolveResult = MAPI_E_NO_SUPPORT;
					book.mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"primary@example.test");
					LPWSTR wide[] = {
						const_cast<LPWSTR>(L"SMTP:primary@example.test"),
						const_cast<LPWSTR>(L"smtp:USER@EXAMPLE.TEST")};
					LPSTR narrow[] = {
						const_cast<LPSTR>("SMTP:primary@example.test"), const_cast<LPSTR>("smtp:USER@EXAMPLE.TEST")};
					book.mailboxProperties.resize(4);
					auto& proxies = book.mailboxProperties[3];
					proxies.ulPropTag = unicode ? PR_EMS_AB_PROXY_ADDRESSES_W : PR_EMS_AB_PROXY_ADDRESSES_A;
					if (unicode)
						proxies.Value.MVszW = {_countof(wide), wide};
					else
						proxies.Value.MVszA = {_countof(narrow), narrow};
					std::wstring dn = L"stale";
					Assert::AreEqual(S_OK, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
					Assert::AreEqual(std::wstring(L"/o=Test/ou=Test/cn=Recipients/cn=User"), dn);
					Assert::AreEqual(size_t{2}, book.requestedLookups.size());
					Assert::AreEqual(std::wstring(L"user@example.test"), book.requestedLookups[0]);
					Assert::AreEqual(std::wstring(L"user"), book.requestedLookups[1]);
					Assert::AreEqual(1UL, book.getPropsCalls);
					Assert::AreEqual(1UL, book.refs());
				}
			}
		}

		TEST_METHOD(VerifiedProxyDoesNotRequirePrimarySMTPProperty)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			ConfigureMailbox(book);
			book.mailboxProperties[2].ulPropTag = CHANGE_PROP_TYPE(PR_SMTP_ADDRESS_W, PT_ERROR);
			book.mailboxProperties[2].Value.err = MAPI_E_NOT_FOUND;
			LPWSTR values[] = {const_cast<LPWSTR>(L"SMTP:user@example.test")};
			book.mailboxProperties.resize(4);
			book.mailboxProperties[3].ulPropTag = PR_EMS_AB_PROXY_ADDRESSES_W;
			book.mailboxProperties[3].Value.MVszW = {_countof(values), values};
			book.getPropsResult = MAPI_W_ERRORS_RETURNED;
			std::wstring dn;
			Assert::AreEqual(S_OK, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::IsFalse(dn.empty());
			Assert::AreEqual(1UL, book.refs());
		}

		TEST_METHOD(RejectsUnverifiedAliasesAndOtherProxyAddressTypes)
		{
			unittest::testMapiMemory memory;
			LPWSTR values[] = {
				nullptr,
				const_cast<LPWSTR>(L""),
				const_cast<LPWSTR>(L"SMTP"),
				const_cast<LPWSTR>(L"X500:user@example.test"),
				const_cast<LPWSTR>(L"SIP:user@example.test"),
				const_cast<LPWSTR>(L"smtp:user@example.test.invalid"),
				const_cast<LPWSTR>(L"smtp:otheruser@example.test")};
			for (const auto malformed : {0, 1, 2})
			{
				TestAddressBook book;
				ConfigureMailbox(book);
				book.mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"primary@example.test");
				book.mailboxProperties.resize(4);
				auto& proxies = book.mailboxProperties[3];
				proxies.ulPropTag = PR_EMS_AB_PROXY_ADDRESSES_W;
				proxies.Value.MVszW = {_countof(values), malformed == 1 ? nullptr : values};
				if (malformed == 2)
				{
					proxies.ulPropTag = CHANGE_PROP_TYPE(PR_EMS_AB_PROXY_ADDRESSES_W, PT_ERROR);
					proxies.Value.err = MAPI_E_NOT_FOUND;
				}

				std::wstring dn = L"stale";
				Assert::AreEqual(MAPI_E_NOT_FOUND, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::IsTrue(dn.empty());
				Assert::AreEqual(1UL, book.refs());
			}
		}

		TEST_METHOD(VerifiedAliasStillRequiresExchangeMailUserAndLegacyDN)
		{
			unittest::testMapiMemory memory;
			for (const auto invalid : {0, 1, 2})
			{
				TestAddressBook book;
				ConfigureMailbox(book);
				book.mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"primary@example.test");
				LPWSTR values[] = {const_cast<LPWSTR>(L"smtp:user@example.test")};
				book.mailboxProperties.resize(4);
				book.mailboxProperties[3].ulPropTag = PR_EMS_AB_PROXY_ADDRESSES_W;
				book.mailboxProperties[3].Value.MVszW = {_countof(values), values};
				if (invalid == 0) book.mailboxProperties[0].Value.lpszW = const_cast<LPWSTR>(L"SMTP");
				if (invalid == 1) book.mailboxProperties[1].Value.lpszW = const_cast<LPWSTR>(L"");
				if (invalid == 2) book.resolvedObjectType = MAPI_DISTLIST;
				std::wstring dn = L"stale";
				Assert::AreEqual(MAPI_E_NOT_FOUND, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::IsTrue(dn.empty());
				Assert::AreEqual(1UL, book.refs());
			}
		}

		TEST_METHOD(RejectsAliasClaimedByDifferentExchangeIdentities)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			TestAddressBook conflict;
			ConfigureMailbox(book);
			ConfigureMailbox(conflict);
			LPWSTR values[] = {const_cast<LPWSTR>(L"smtp:user@example.test")};
			for (const auto recipient : {&book, &conflict})
			{
				recipient->mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"primary@example.test");
				recipient->mailboxProperties.resize(4);
				recipient->mailboxProperties[3].ulPropTag = PR_EMS_AB_PROXY_ADDRESSES_W;
				recipient->mailboxProperties[3].Value.MVszW = {_countof(values), values};
			}

			conflict.mailboxProperties[1].Value.lpszW = const_cast<LPWSTR>(L"/o=Test/ou=Test/cn=Recipients/cn=Other");
			book.searchContainers = {&book, &conflict};
			std::wstring dn = L"stale";
			Assert::AreEqual(MAPI_E_AMBIGUOUS_RECIP, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::IsTrue(dn.empty());
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(1UL, conflict.refs());
		}

		TEST_METHOD(SearchesPastContactsAndNearMatchesForVerifiedExchangeMailbox)
		{
			unittest::testMapiMemory memory;
			for (const auto contact : {false, true})
			{
				TestAddressBook book;
				TestAddressBook exchange;
				ConfigureMailbox(book);
				ConfigureMailbox(exchange);
				if (contact)
					book.mailboxProperties[0].Value.lpszW = const_cast<LPWSTR>(L"SMTP");
				else
					book.mailboxProperties[2].Value.lpszW = const_cast<LPWSTR>(L"other@example.test");
				book.searchContainers = {&book, &exchange};
				std::wstring dn = L"stale";
				Assert::AreEqual(S_OK, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::AreEqual(std::wstring(L"/o=Test/ou=Test/cn=Recipients/cn=User"), dn);
				Assert::AreEqual(2UL, book.containerOpenCalls);
				Assert::AreEqual(2UL, book.resolveCalls);
				Assert::AreEqual(1UL, exchange.resolveCalls);
				Assert::AreEqual(1UL, book.refs());
				Assert::AreEqual(1UL, exchange.refs());
			}
		}

		TEST_METHOD(ContinuesPastUnresolvedContainerWithoutOpeningItsRecipient)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			TestAddressBook exchange;
			ConfigureMailbox(exchange);
			book.resolutionStatus = MAPI_UNRESOLVED;
			book.searchContainers = {&book, &exchange};
			std::wstring dn;
			Assert::AreEqual(S_OK, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::AreEqual(0UL, book.openCalls);
			Assert::AreEqual(1UL, exchange.openCalls);
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(1UL, exchange.refs());
		}

		TEST_METHOD(DeduplicatesSameExchangeIdentityAcrossContainers)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			TestAddressBook duplicate;
			ConfigureMailbox(book);
			ConfigureMailbox(duplicate);
			duplicate.mailboxProperties[1].Value.lpszW = const_cast<LPWSTR>(L"/o=test/ou=test/cn=recipients/cn=user");
			book.searchContainers = {&book, &duplicate};
			std::wstring dn;
			Assert::AreEqual(S_OK, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::AreEqual(1UL, book.getPropsCalls);
			Assert::AreEqual(1UL, duplicate.getPropsCalls);
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(1UL, duplicate.refs());
		}

		TEST_METHOD(RejectsConflictingExchangeIdentitiesWithoutOpeningAnyStore)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			TestAddressBook conflict;
			ConfigureMailbox(book);
			ConfigureMailbox(conflict);
			conflict.mailboxProperties[1].Value.lpszW = const_cast<LPWSTR>(L"/o=Test/ou=Test/cn=Recipients/cn=Other");
			book.searchContainers = {&book, &conflict};
			std::wstring dn = L"stale";
			Assert::AreEqual(MAPI_E_AMBIGUOUS_RECIP, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::IsTrue(dn.empty());
			TestSession session;
			session.addressBook = &book;
			LPMDB mailbox{};
			Assert::AreEqual(
				MAPI_E_AMBIGUOUS_RECIP, mapi::store::OpenMailboxBySMTP(&session, L"user@example.test", &mailbox));
			Assert::IsNull(mailbox);
			Assert::AreEqual(0UL, session.storeTableCalls);
			Assert::AreEqual(0UL, session.storeOpenCalls);
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(1UL, conflict.refs());
		}

		TEST_METHOD(RejectsAmbiguousContainerResolutionStatus)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			book.resolutionStatus = MAPI_AMBIGUOUS;
			std::wstring dn = L"stale";
			Assert::AreEqual(MAPI_E_AMBIGUOUS_RECIP, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::IsTrue(dn.empty());
			Assert::AreEqual(0UL, book.openCalls);
			Assert::AreEqual(1UL, book.refs());
		}

		TEST_METHOD(DoesNotHideLaterProviderFailureBehindEarlierMatch)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			TestAddressBook failed;
			ConfigureMailbox(book);
			failed.resolveResult = MAPI_E_NO_ACCESS;
			book.searchContainers = {&book, &failed};
			std::wstring dn = L"stale";
			Assert::AreEqual(MAPI_E_NO_ACCESS, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::IsTrue(dn.empty());
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(1UL, failed.refs());
		}

		TEST_METHOD(PreservesSearchPathAndContainerOpenErrors)
		{
			unittest::testMapiMemory memory;
			for (const auto failSearch : {false, true})
			{
				TestAddressBook book;
				if (failSearch)
					book.searchPathResult = MAPI_E_NO_ACCESS;
				else
					book.containerOpenResult = MAPI_E_NOT_FOUND;
				std::wstring dn = L"stale";
				Assert::AreEqual(
					failSearch ? MAPI_E_NO_ACCESS : MAPI_E_NOT_FOUND,
					mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::IsTrue(dn.empty());
				Assert::AreEqual(0UL, book.resolveCalls);
				Assert::AreEqual(1UL, book.refs());
			}
		}

		TEST_METHOD(RejectsMissingAndMalformedSearchPaths)
		{
			unittest::testMapiMemory memory;
			for (const auto missing : {false, true})
			{
				TestAddressBook book;
				if (missing)
					book.returnSearchPath = false;
				else
					book.searchEntryIDTag = CHANGE_PROP_TYPE(PR_ENTRYID, PT_ERROR);
				std::wstring dn = L"stale";
				Assert::AreEqual(
					MAPI_E_INVALID_OBJECT, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::IsTrue(dn.empty());
				Assert::AreEqual(0UL, book.containerOpenCalls);
			}

			TestAddressBook empty;
			empty.searchContainers.clear();
			std::wstring dn = L"stale";
			Assert::AreEqual(MAPI_E_NOT_FOUND, mapi::ab::ResolveExchangeMailbox(&empty, L"user@example.test", dn));
			Assert::IsTrue(dn.empty());
		}

		TEST_METHOD(SearchesContentsWhenContainerDoesNotSupportResolveNames)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			ConfigureMailbox(book);
			book.resolveResult = MAPI_E_NO_SUPPORT;
			std::wstring dn;
			Assert::AreEqual(S_OK, mapi::ab::ResolveExchangeMailbox(&book, L"User@Example.Test", dn));
			Assert::AreEqual(std::wstring(L"/o=Test/ou=Test/cn=Recipients/cn=User"), dn);
			Assert::AreEqual(1UL, book.contentsTableCalls);
			Assert::AreEqual(2UL, book.queryCalls);
			Assert::AreEqual(1UL, book.refs());
		}

		TEST_METHOD(ContentsSearchVerifiesEveryCandidateAcrossBatches)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			TestAddressBook contact;
			ConfigureMailbox(book);
			ConfigureMailbox(contact);
			contact.mailboxProperties[0].Value.lpszW = const_cast<LPWSTR>(L"SMTP");
			book.resolveResult = MAPI_E_NO_SUPPORT;
			book.tableRecipients.assign(20, &contact);
			book.tableRecipients.push_back(&book);
			std::wstring dn;
			Assert::AreEqual(S_OK, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::AreEqual(3UL, book.queryCalls);
			Assert::AreEqual(20UL, contact.getPropsCalls);
			Assert::AreEqual(1UL, book.getPropsCalls);
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(1UL, contact.refs());
		}

		TEST_METHOD(ContentsSearchRejectsConflictingExactMatches)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			TestAddressBook conflict;
			ConfigureMailbox(book);
			ConfigureMailbox(conflict);
			conflict.mailboxProperties[1].Value.lpszW = const_cast<LPWSTR>(L"/o=Test/ou=Test/cn=Recipients/cn=Other");
			book.resolveResult = MAPI_E_NO_SUPPORT;
			book.tableRecipients = {&book, &conflict};
			std::wstring dn = L"stale";
			Assert::AreEqual(MAPI_E_AMBIGUOUS_RECIP, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::IsTrue(dn.empty());
			Assert::AreEqual(1UL, book.refs());
			Assert::AreEqual(1UL, conflict.refs());
		}

		TEST_METHOD(ContentsSearchPreservesTableFailuresAndCleansUp)
		{
			unittest::testMapiMemory memory;
			for (const auto stage : {0, 1, 2, 3})
			{
				TestAddressBook book;
				ConfigureMailbox(book);
				book.resolveResult = MAPI_E_NO_SUPPORT;
				if (stage == 0) book.contentsTableResult = MAPI_E_NO_ACCESS;
				if (stage == 1) book.columnsResult = MAPI_E_NO_ACCESS;
				if (stage == 2) book.restrictResult = MAPI_E_NO_ACCESS;
				if (stage == 3) book.queryResult = MAPI_E_NO_ACCESS;
				std::wstring dn = L"stale";
				Assert::AreEqual(MAPI_E_NO_ACCESS, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
				Assert::IsTrue(dn.empty());
				Assert::AreEqual(0UL, book.getPropsCalls);
				Assert::AreEqual(1UL, book.refs());
			}
		}

		TEST_METHOD(DoesNotUseContentsFallbackForRecipientPropertyErrors)
		{
			unittest::testMapiMemory memory;
			TestAddressBook book;
			ConfigureMailbox(book);
			book.getPropsResult = MAPI_E_NO_SUPPORT;
			std::wstring dn = L"stale";
			Assert::AreEqual(MAPI_E_NO_SUPPORT, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
			Assert::IsTrue(dn.empty());
			Assert::AreEqual(0UL, book.contentsTableCalls);
			Assert::AreEqual(1UL, book.refs());
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
				Assert::AreEqual(MAPI_E_NOT_FOUND, mapi::ab::ResolveExchangeMailbox(&book, L"user@example.test", dn));
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
		std::function<std::wstring()> m_originalServerPrompt;
	};
} // namespace mapifunctionstest
