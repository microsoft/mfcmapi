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
			const ULONG aulOneOffIDs[] = {dispidFormStorage,
										  dispidPageDirStream,
										  dispidFormPropStream,
										  dispidScriptStream,
										  dispidPropDefStream,
										  dispidCustomFlag};
			constexpr ULONG ulNumOneOffIDs = _countof(aulOneOffIDs);

			const auto formStorageID = MAPINAMEID{const_cast<LPGUID>(&guid::PSETID_Common), MNID_ID, dispidFormStorage};
			const auto pageDirStreamID =
				MAPINAMEID{const_cast<LPGUID>(&guid::PSETID_Common), MNID_ID, dispidPageDirStream};
			const auto formPropStreamID =
				MAPINAMEID{const_cast<LPGUID>(&guid::PSETID_Common), MNID_ID, dispidFormPropStream};

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

			Assert::AreEqual(true, formStorage1.match(formStorage1, false, false, false));

			// Should fail
			Assert::AreEqual(false, formStorage1.match(formStorage2, true, true, true));

			// Should all work
			Assert::AreEqual(true, formStorage1.match(formStorage2, false, true, true));
			Assert::AreEqual(true, formStorage1.match(formStorage2, false, false, true));
			Assert::AreEqual(true, formStorage1.match(formStorage2, false, true, false));
			Assert::AreEqual(true, formStorage1.match(formStorage2, false, false, false));
		}
	};
} // namespace namedproptest