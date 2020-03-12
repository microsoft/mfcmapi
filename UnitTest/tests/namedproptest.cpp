#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/mapi/extraPropTags.h>

namespace namedproptest
{
	TEST_CLASS(namedproptest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_Match)
		{
			const auto sig1 = std::vector<BYTE>({1, 2, 3, 4});
			const auto sig2 = std::vector<BYTE>({5, 6, 7, 8, 9});

			const auto formStorageID = MAPINAMEID{const_cast<LPGUID>(&guid::PSETID_Common), MNID_ID, dispidFormStorage};
			const auto formStorageName = MAPINAMEID{const_cast<LPGUID>(&guid::PSETID_Common), MNID_ID, {.lpwstrName = L"name"}};
			const auto pageDirStreamID =
				MAPINAMEID{const_cast<LPGUID>(&guid::PSETID_Common), MNID_ID, dispidPageDirStream};

			const auto formStorage1 = cache::namedPropCacheEntry(&formStorageID, 0x1111, sig1);
			const auto formStorage2 = cache::namedPropCacheEntry(&formStorageID, 0x1111, sig2);

			// Test all forms of match
			Assert::AreEqual(true, formStorage1.match(formStorage1, true, true, true));
			Assert::AreEqual(true, formStorage1.match(formStorage1, false, true, true));
			Assert::AreEqual(true, formStorage1.match(formStorage1, true, false, true));
			Assert::AreEqual(true, formStorage1.match(formStorage1, true, true, false));
			Assert::AreEqual(true, formStorage1.match(formStorage1, true, false, false));
			Assert::AreEqual(true, formStorage1.match(formStorage1, false, true, false));
			Assert::AreEqual(true, formStorage1.match(formStorage1, false, false, true));
			Assert::AreEqual(true, formStorage1.match(formStorage1, false, false, false));

			// Should fail
			Assert::AreEqual(false, formStorage1.match(formStorage2, true, true, true));

			// Odd comparisons
			Assert::AreEqual(
				true, formStorage1.match(cache::namedPropCacheEntry(&formStorageID, 0x1111, sig1), true, true, true));
			Assert::AreEqual(
				true, formStorage1.match(cache::namedPropCacheEntry(&formStorageID, 0x1112, sig1), true, false, true));
			Assert::AreEqual(
				true,
				formStorage1.match(cache::namedPropCacheEntry(&pageDirStreamID, 0x1111, sig1), true, true, false));
			Assert::AreEqual(
				true,
				formStorage1.match(cache::namedPropCacheEntry(&formStorageName, 0x1111, sig1), true, true, false));
			Assert::AreEqual(
				false,
				formStorage1.match(cache::namedPropCacheEntry(&formStorageName, 0x1111, sig1), true, true, true));

			// Should all work
			Assert::AreEqual(true, formStorage1.match(formStorage2, false, true, true));
			Assert::AreEqual(true, formStorage1.match(formStorage2, false, false, true));
			Assert::AreEqual(true, formStorage1.match(formStorage2, false, true, false));
			Assert::AreEqual(true, formStorage1.match(formStorage2, false, false, false));

			// Compare given a signature, MAPINAMEID
			// _Check_return_ bool match(_In_ const std::vector<BYTE>& _sig, _In_ const MAPINAMEID& _mapiNameId) const;
			Assert::AreEqual(true, formStorage1.match(sig1, formStorageID));
			Assert::AreEqual(false, formStorage1.match(sig2, formStorageID));
			Assert::AreEqual(false, formStorage1.match(sig1, pageDirStreamID));

			// Compare given a signature and property ID (ulPropID)
			// _Check_return_ bool match(_In_ const std::vector<BYTE>& _sig, ULONG _ulPropID) const;
			Assert::AreEqual(true, formStorage1.match(sig1, 0x1111));
			Assert::AreEqual(false, formStorage1.match(sig1, 0x1112));
			Assert::AreEqual(false, formStorage1.match(sig2, 0x1111));

			// Compare given a id, MAPINAMEID
			// _Check_return_ bool match(ULONG _ulPropID, _In_ const MAPINAMEID& _mapiNameId) const noexcept;
			Assert::AreEqual(true, formStorage1.match(0x1111, formStorageID));
			Assert::AreEqual(false, formStorage1.match(0x1112, formStorageID));
			Assert::AreEqual(false, formStorage1.match(0x1111, pageDirStreamID));
		}
	};
} // namespace namedproptest