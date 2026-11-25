#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <UI/Controls/SortList/TableLoad.h>
#include <core/utility/error.h>
#include <CommCtrl.h>

namespace sortlisttest
{
	LRESULT CALLBACK FailInsert(HWND window, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR)
	{
		if (message == LVM_INSERTITEMW) return -1;
		return DefSubclassProc(window, message, wParam, lParam);
	}

	TEST_CLASS(sortlisttest)
	{
	public:
		static const bool dummy_var = true;
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(InsertionFailureDoesNotReturnItemData)
		{
			INITCOMMONCONTROLSEX commonControls{sizeof(commonControls), ICC_LISTVIEW_CLASSES};
			Assert::AreEqual(S_OK, WC_H(InitCommonControlsEx(&commonControls) ? S_OK : E_FAIL));
			const auto instance = WC_D(HMODULE, GetModuleHandleW(nullptr));
			Assert::IsNotNull(instance);
			const auto window = WC_D(
				HWND,
				CreateWindowExW(
					0, WC_LISTVIEWW, L"", WS_POPUP | LVS_REPORT, 0, 0, 100, 100, nullptr, nullptr, instance, nullptr));
			Assert::IsNotNull(window);
			struct Cleanup
			{
				HWND window;
				~Cleanup() { WC_B_S(DestroyWindow(window)); }
			} cleanup{window};
			Assert::AreEqual(S_OK, WC_H(SetWindowSubclass(window, FailInsert, 1, 0) ? S_OK : E_FAIL));
			Assert::IsNull(controls::sortlistctrl::InsertSortListRow(window, 0, L"failed", 0, 0));
			Assert::AreEqual(0, ListView_GetItemCount(window));
			Assert::AreEqual(S_OK, WC_H(RemoveWindowSubclass(window, FailInsert, 1) ? S_OK : E_FAIL));

			auto data = std::unique_ptr<sortlistdata::sortListData>(
				controls::sortlistctrl::InsertSortListRow(window, 0, L"success", 0, 0));
			Assert::IsNotNull(data.get());
			Assert::AreEqual(1, ListView_GetItemCount(window));
			LVITEMW item{};
			item.mask = LVIF_PARAM;
			item.iItem = 0;
			Assert::AreEqual(S_OK, WC_H(ListView_GetItem(window, &item) ? S_OK : E_FAIL));
			Assert::IsTrue(item.lParam == reinterpret_cast<LPARAM>(data.get()));
			Assert::AreEqual(S_OK, WC_H(ListView_DeleteItem(window, 0) ? S_OK : E_FAIL));
		}
	};
} // namespace sortlisttest
