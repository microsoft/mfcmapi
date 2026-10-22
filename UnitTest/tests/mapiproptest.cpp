#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/mapiFunctions.h>
#include <core/propertyBag/mapiPropPropertyBag.h>
#include <core/utility/error.h>
#include <memory>

namespace mapiproptest
{
	class TestProperty final : public IMAPIProp
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
		HRESULT STDMETHODCALLTYPE SaveChanges(ULONG flags) override
		{
			++saves;
			saveFlags = flags;
			return saveResult;
		}
		HRESULT STDMETHODCALLTYPE GetProps(LPSPropTagArray, ULONG, ULONG*, LPSPropValue*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE GetPropList(ULONG flags, LPSPropTagArray* tags) override
		{
			++lists;
			listFlags = flags;
			*tags = nullptr;
			if (FAILED(listResult) || nullList) return listResult;
			const auto allocated = MAPIAllocateBuffer(
				static_cast<ULONG>(CbNewSPropTagArray(listed.size())), reinterpret_cast<void**>(tags));
			if (FAILED(allocated)) return allocated;
			(*tags)->cValues = static_cast<ULONG>(listed.size());
			std::copy(listed.begin(), listed.end(), (*tags)->aulPropTag);
			return listResult;
		}
		HRESULT STDMETHODCALLTYPE OpenProperty(ULONG, LPCIID, ULONG, ULONG, LPUNKNOWN*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE SetProps(ULONG, LPSPropValue, LPSPropProblemArray*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE DeleteProps(LPSPropTagArray tags, LPSPropProblemArray* problems) override
		{
			*problems = nullptr;
			if (!tags || tags->cValues != 1) return MAPI_E_INVALID_PARAMETER;
			const auto tag = tags->aulPropTag[0];
			deleted.push_back(tag);
			if (FAILED(deleteResult)) return deleteResult;
			if (rejectUntyped && PROP_TYPE(tag) == PT_UNSPECIFIED) return MAPI_E_NOT_FOUND;
			if (FAILED(problem))
			{
				const auto allocated =
					MAPIAllocateBuffer(sizeof(SPropProblemArray), reinterpret_cast<void**>(problems));
				if (FAILED(allocated)) return allocated;
				(*problems)->cProblem = 1;
				(*problems)->aProblem[0] = {0, tag, problem};
			}
			else if (tag == stored)
			{
				present = false;
			}
			return S_OK;
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

		ULONG stored{0x00070102};
		std::vector<ULONG> listed{PR_SUBJECT_W, stored};
		std::vector<ULONG> deleted;
		HRESULT listResult{S_OK};
		HRESULT deleteResult{S_OK};
		HRESULT problem{S_OK};
		HRESULT saveResult{S_OK};
		ULONG lists{};
		ULONG listFlags{};
		ULONG saves{};
		ULONG saveFlags{};
		bool present{true};
		bool nullList{};
		bool rejectUntyped{};

	private:
		ULONG m_refs{1};
	};

	TEST_CLASS(mapiproptest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD_INITIALIZE(captureErrors)
		{
			m_memory = std::make_unique<unittest::testMapiMemory>();
			m_errors.clear();
			m_originalDisplayError = std::move(error::displayError);
			error::displayError = [this](const std::wstring& text) { m_errors.push_back(text); };
		}

		TEST_METHOD_CLEANUP(restoreErrors)
		{
			error::displayError = std::move(m_originalDisplayError);
			m_memory.reset();
		}

		TEST_METHOD(DeletesErrorValuedPropertyWithItsStoredType)
		{
			for (const ULONG tag : {0x00070102UL, 0x0007001FUL, 0x0007001EUL, 0x00070003UL, 0x800A1102UL})
			{
				TestProperty prop;
				prop.stored = tag;
				prop.listed = {PR_SUBJECT_W, tag, PR_DISPLAY_NAME_W};
				Assert::AreEqual(S_OK, mapi::DeleteProperty(&prop, CHANGE_PROP_TYPE(tag, PT_ERROR)));
				Assert::IsFalse(prop.present);
				Assert::AreEqual(size_t{1}, prop.deleted.size());
				Assert::AreEqual(tag, prop.deleted[0]);
				Assert::AreEqual(1UL, prop.lists);
				Assert::AreEqual(0UL, prop.listFlags);
				Assert::AreEqual(1UL, prop.saves);
				Assert::AreEqual(static_cast<ULONG>(KEEP_OPEN_READWRITE), prop.saveFlags);
			}
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(PropertyBagDeletesErrorValuedPropertyWithoutGuessing)
		{
			TestProperty prop;
			propertybag::mapiPropPropertyBag bag(&prop, nullptr, false);
			Assert::AreEqual(S_OK, bag.DeleteProp(CHANGE_PROP_TYPE(prop.stored, PT_ERROR), L""));
			Assert::IsFalse(prop.present);
			Assert::AreEqual(size_t{1}, prop.deleted.size());
			Assert::AreEqual(prop.stored, prop.deleted[0]);
		}

		TEST_METHOD(TypedAndExplicitlyUnspecifiedDeletesDoNotEnumerate)
		{
			for (const ULONG tag : {0x00070102UL, 0x00070000UL})
			{
				TestProperty prop;
				Assert::AreEqual(S_OK, mapi::DeleteProperty(&prop, tag));
				Assert::AreEqual(0UL, prop.lists);
				Assert::AreEqual(tag, prop.deleted[0]);
				Assert::AreEqual(1UL, prop.saves);
			}
		}

		TEST_METHOD(UnavailablePropertyListsPreserveUntypedFallback)
		{
			for (const auto result : {MAPI_E_NO_SUPPORT, MAPI_E_CALL_FAILED})
			{
				TestProperty prop;
				prop.listResult = result;
				Assert::AreEqual(S_OK, mapi::DeleteProperty(&prop, CHANGE_PROP_TYPE(prop.stored, PT_ERROR)));
				Assert::AreEqual(1UL, prop.lists);
				Assert::AreEqual(CHANGE_PROP_TYPE(prop.stored, PT_UNSPECIFIED), prop.deleted[0]);
			}
		}

		TEST_METHOD(MissingOrUnusableListedTypesPreserveUntypedFallback)
		{
			for (const ULONG tag : {PR_SUBJECT_W, 0x0007000AUL, 0x00070000UL})
			{
				TestProperty prop;
				prop.listed = {tag};
				Assert::AreEqual(S_OK, mapi::DeleteProperty(&prop, CHANGE_PROP_TYPE(prop.stored, PT_ERROR)));
				Assert::AreEqual(CHANGE_PROP_TYPE(prop.stored, PT_UNSPECIFIED), prop.deleted[0]);
			}
			TestProperty prop;
			prop.nullList = true;
			Assert::AreEqual(S_OK, mapi::DeleteProperty(&prop, CHANGE_PROP_TYPE(prop.stored, PT_ERROR)));
			Assert::AreEqual(CHANGE_PROP_TYPE(prop.stored, PT_UNSPECIFIED), prop.deleted[0]);
		}

		TEST_METHOD(PropertyBagRetainsStreamFallbackWhenEnumerationIsUnsupported)
		{
			TestProperty prop;
			prop.listResult = MAPI_E_NO_SUPPORT;
			prop.rejectUntyped = true;
			propertybag::mapiPropPropertyBag bag(&prop, nullptr, false);
			Assert::AreEqual(S_OK, bag.DeleteProp(CHANGE_PROP_TYPE(prop.stored, PT_ERROR), L""));
			Assert::IsFalse(prop.present);
			Assert::AreEqual(size_t{2}, prop.deleted.size());
			Assert::AreEqual(CHANGE_PROP_TYPE(prop.stored, PT_UNSPECIFIED), prop.deleted[0]);
			Assert::AreEqual(prop.stored, prop.deleted[1]);
		}

		TEST_METHOD(DeleteFailureIsReturnedWithoutSaving)
		{
			TestProperty prop;
			prop.deleteResult = MAPI_E_NO_ACCESS;
			Assert::AreEqual(MAPI_E_NO_ACCESS, mapi::DeleteProperty(&prop, CHANGE_PROP_TYPE(prop.stored, PT_ERROR)));
			Assert::AreEqual(0UL, prop.saves);
			Assert::IsTrue(prop.present);
			Assert::IsFalse(m_errors.empty());
		}

		TEST_METHOD(ProblemArrayFailureIsReturnedWithoutSaving)
		{
			TestProperty prop;
			prop.problem = MAPI_E_NO_ACCESS;
			Assert::AreEqual(MAPI_E_NO_ACCESS, mapi::DeleteProperty(&prop, CHANGE_PROP_TYPE(prop.stored, PT_ERROR)));
			Assert::AreEqual(0UL, prop.saves);
			Assert::IsTrue(prop.present);
		}

		TEST_METHOD(SaveFailureIsReturned)
		{
			TestProperty prop;
			prop.saveResult = MAPI_E_NO_ACCESS;
			Assert::AreEqual(MAPI_E_NO_ACCESS, mapi::DeleteProperty(&prop, CHANGE_PROP_TYPE(prop.stored, PT_ERROR)));
			Assert::AreEqual(1UL, prop.saves);
			Assert::IsFalse(m_errors.empty());
		}

		TEST_METHOD(NullPropertyIsRejected)
		{
			Assert::AreEqual(MAPI_E_INVALID_PARAMETER, mapi::DeleteProperty(nullptr, 0x0007000A));
		}

	private:
		std::unique_ptr<unittest::testMapiMemory> m_memory;
		std::function<void(const std::wstring&)> m_originalDisplayError;
		std::vector<std::wstring> m_errors;
	};
} // namespace mapiproptest
