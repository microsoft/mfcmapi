#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/cache/globalCache.h>
#include <core/propertyBag/rowPropertyBag.h>
#include <memory>

namespace globalcachetest
{
	struct ReferenceState
	{
		ULONG count{1};
		bool destroyed{};
	};

	class TestMessage final : public IMessage
	{
	public:
		std::shared_ptr<ReferenceState> state = std::make_shared<ReferenceState>();

		STDMETHODIMP QueryInterface(REFIID iid, void** object) override
		{
			if (!object) return E_POINTER;
			*object = nullptr;
			if (iid != IID_IUnknown && iid != IID_IMAPIProp && iid != IID_IMessage) return E_NOINTERFACE;
			*object = static_cast<IMessage*>(this);
			AddRef();
			return S_OK;
		}

		STDMETHODIMP_(ULONG) AddRef() override { return ++state->count; }
		STDMETHODIMP_(ULONG) Release() override
		{
			const auto count = --state->count;
			if (!count) delete this;
			return count;
		}

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
		STDMETHODIMP GetAttachmentTable(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP OpenAttach(ULONG, LPCIID, ULONG, LPATTACH*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP CreateAttach(LPCIID, ULONG, ULONG*, LPATTACH*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP DeleteAttach(ULONG, ULONG_PTR, LPMAPIPROGRESS, ULONG) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP GetRecipientTable(ULONG, LPMAPITABLE*) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP ModifyRecipients(ULONG, LPADRLIST) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP SubmitMessage(ULONG) override { return MAPI_E_NO_SUPPORT; }
		STDMETHODIMP SetReadFlag(ULONG) override { return MAPI_E_NO_SUPPORT; }

	private:
		~TestMessage() { state->destroyed = true; }
	};

	struct ReleaseMessage
	{
		void operator()(TestMessage* message) const { message->Release(); }
	};
	using MessageOwner = std::unique_ptr<TestMessage, ReleaseMessage>;

	class TestPropertyBag : public propertybag::rowPropertyBag
	{
	public:
		explicit TestPropertyBag(IMAPIProp* object) : rowPropertyBag(nullptr, false), m_object(object)
		{
			if (m_object) m_object->AddRef();
		}

		~TestPropertyBag() override
		{
			if (m_object) m_object->Release();
		}

		LPMAPIPROP GetMAPIProp() const override { return m_object; }

	private:
		IMAPIProp* m_object;
	};

	std::shared_ptr<sortlistdata::propModelData> CopiedProperty()
	{
		auto model = std::make_shared<model::mapiRowModel>();
		model->ulPropTag(PROP_TAG(PT_BINARY, 0x7FFF));
		auto prop = std::make_shared<sortlistdata::propModelData>(model);
		// Empty display text prevents the real cache from changing the Windows clipboard.
		Assert::IsTrue(prop->ToString().empty());
		return prop;
	}

	TEST_CLASS(globalCacheTest)
	{
	public:
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }
		TEST_METHOD_INITIALIZE(reset) { cache::CGlobalCache::getInstance().SetABEntriesToCopy(nullptr); }
		TEST_METHOD_CLEANUP(cleanup) { cache::CGlobalCache::getInstance().SetABEntriesToCopy(nullptr); }

		TEST_METHOD(PropertyCopyExposesSourceObject)
		{
			auto& cache = cache::CGlobalCache::getInstance();
			const auto source = MessageOwner(new TestMessage());
			const auto bag = std::make_shared<TestPropertyBag>(source.get());
			const auto prop = CopiedProperty();
			cache.SetPropertyToCopy(prop, bag);

			Assert::AreEqual<ULONG>(BUFFER_PROPTAG | BUFFER_SOURCEPROPOBJ, cache.GetBufferStatus());
			Assert::IsTrue(cache.GetPropertyToCopy() == prop);
			Assert::IsTrue(cache.GetSourcePropBag() == bag);
			const auto object = cache.GetSourcePropObject();
			const auto same = object == source.get();
			if (object) object->Release();
			Assert::IsTrue(same);
		}

		TEST_METHOD(PropertyCopyOwnsAnIndependentSourceReference)
		{
			auto& cache = cache::CGlobalCache::getInstance();
			const auto source = MessageOwner(new TestMessage());
			const auto bag = std::make_shared<TestPropertyBag>(source.get());
			Assert::AreEqual<ULONG>(2, source->state->count);
			cache.SetPropertyToCopy(CopiedProperty(), bag);
			Assert::AreEqual<ULONG>(3, source->state->count);
			cache.SetABEntriesToCopy(nullptr);
			Assert::AreEqual<ULONG>(2, source->state->count);
		}

		TEST_METHOD(SourceGetterReturnsAnOwnedReference)
		{
			auto& cache = cache::CGlobalCache::getInstance();
			const auto source = MessageOwner(new TestMessage());
			cache.SetPropertyToCopy(CopiedProperty(), std::make_shared<TestPropertyBag>(source.get()));
			const auto before = source->state->count;
			const auto object = cache.GetSourcePropObject();
			const auto acquired = source->state->count;
			if (object) object->Release();
			Assert::IsNotNull(object);
			Assert::AreEqual<ULONG>(before + 1, acquired);
			Assert::AreEqual<ULONG>(before, source->state->count);
		}

		TEST_METHOD(CachedSourceSurvivesItsOriginalOwners)
		{
			auto& cache = cache::CGlobalCache::getInstance();
			auto source = MessageOwner(new TestMessage());
			const auto state = source->state;
			auto bag = std::make_shared<TestPropertyBag>(source.get());
			const std::weak_ptr<TestPropertyBag> weakBag = bag;
			cache.SetPropertyToCopy(CopiedProperty(), bag);
			bag.reset();
			source.reset();
			Assert::IsFalse(weakBag.expired());
			Assert::IsFalse(state->destroyed);
			Assert::AreEqual<ULONG>(2, state->count);

			cache.SetABEntriesToCopy(nullptr);
			Assert::IsTrue(weakBag.expired());
			Assert::IsTrue(state->destroyed);
			Assert::AreEqual<ULONG>(0, state->count);
			Assert::AreEqual<ULONG>(BUFFER_EMPTY, cache.GetBufferStatus());
			Assert::IsNull(cache.GetSourcePropObject());
			Assert::IsTrue(cache.GetSourcePropBag() == nullptr);
			Assert::IsTrue(cache.GetPropertyToCopy() == nullptr);
		}

		TEST_METHOD(ReplacingPropertyCopyReleasesThePreviousSource)
		{
			auto& cache = cache::CGlobalCache::getInstance();
			const auto first = MessageOwner(new TestMessage());
			const auto second = MessageOwner(new TestMessage());
			cache.SetPropertyToCopy(CopiedProperty(), std::make_shared<TestPropertyBag>(first.get()));
			cache.SetPropertyToCopy(CopiedProperty(), std::make_shared<TestPropertyBag>(second.get()));
			Assert::AreEqual<ULONG>(1, first->state->count);
			Assert::AreEqual<ULONG>(3, second->state->count);
			cache.SetABEntriesToCopy(nullptr);
			Assert::AreEqual<ULONG>(1, second->state->count);
		}

		TEST_METHOD(RowOnlySourceDoesNotAdvertiseAMAPIObject)
		{
			auto& cache = cache::CGlobalCache::getInstance();
			const auto bag = std::make_shared<TestPropertyBag>(nullptr);
			cache.SetPropertyToCopy(CopiedProperty(), bag);
			Assert::AreEqual<ULONG>(BUFFER_PROPTAG, cache.GetBufferStatus());
			Assert::IsNull(cache.GetSourcePropObject());
			Assert::IsTrue(cache.GetSourcePropBag() == bag);
		}

		TEST_METHOD(RowOnlyCopyClearsThePreviousMAPIObject)
		{
			auto& cache = cache::CGlobalCache::getInstance();
			const auto source = MessageOwner(new TestMessage());
			cache.SetPropertyToCopy(CopiedProperty(), std::make_shared<TestPropertyBag>(source.get()));
			cache.SetPropertyToCopy(CopiedProperty(), std::make_shared<TestPropertyBag>(nullptr));
			Assert::AreEqual<ULONG>(1, source->state->count);
			Assert::AreEqual<ULONG>(BUFFER_PROPTAG, cache.GetBufferStatus());
			Assert::IsNull(cache.GetSourcePropObject());
		}

		TEST_METHOD(AttachmentCopyReplacesPropertyCopy)
		{
			auto& cache = cache::CGlobalCache::getInstance();
			const auto propertySource = MessageOwner(new TestMessage());
			const auto message = MessageOwner(new TestMessage());
			cache.SetPropertyToCopy(CopiedProperty(), std::make_shared<TestPropertyBag>(propertySource.get()));
			cache.SetAttachmentsToCopy(message.get(), {1, 3});
			Assert::AreEqual<ULONG>(1, propertySource->state->count);
			Assert::AreEqual<ULONG>(BUFFER_ATTACHMENTS | BUFFER_SOURCEPROPOBJ, cache.GetBufferStatus());
			Assert::IsTrue(cache.GetPropertyToCopy() == nullptr);
			Assert::IsTrue(cache.GetSourcePropBag() == nullptr);
			Assert::IsTrue(cache.GetAttachmentsToCopy() == std::vector<ULONG>({1, 3}));
			const auto object = cache.GetSourcePropObject();
			const auto same = object == message.get();
			if (object) object->Release();
			Assert::IsTrue(same);
			cache.SetABEntriesToCopy(nullptr);
			Assert::AreEqual<ULONG>(1, message->state->count);
		}

		TEST_METHOD(PropertyCopyReplacesAttachmentCopy)
		{
			auto& cache = cache::CGlobalCache::getInstance();
			const auto message = MessageOwner(new TestMessage());
			const auto propertySource = MessageOwner(new TestMessage());
			cache.SetAttachmentsToCopy(message.get(), {2});
			cache.SetPropertyToCopy(CopiedProperty(), std::make_shared<TestPropertyBag>(propertySource.get()));
			Assert::AreEqual<ULONG>(1, message->state->count);
			Assert::AreEqual<ULONG>(3, propertySource->state->count);
			Assert::AreEqual<ULONG>(BUFFER_PROPTAG | BUFFER_SOURCEPROPOBJ, cache.GetBufferStatus());
			Assert::IsTrue(cache.GetAttachmentsToCopy().empty());
		}
	};
} // namespace globalcachetest
