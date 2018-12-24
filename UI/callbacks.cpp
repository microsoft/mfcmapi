#include <StdAfx.h>
#include <UI/callbacks.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <Interpret/Guids.h>
#include <MAPI/MAPIProgress.h>
#include <MAPI/MAPIFunctions.h>

namespace ui
{
	namespace callbacks
	{
		// Takes a tag array (and optional MAPIProp) and displays UI prompting to build an exclusion array
		// Must be freed with MAPIFreeBuffer
		LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB)
		{
			dialog::editor::CTagArrayEditor TagEditor(
				nullptr, IDS_TAGSTOEXCLUDE, IDS_TAGSTOEXCLUDEPROMPT, nullptr, lpTagArray, bIsAB, lpProp);
			if (TagEditor.DisplayDialog())
			{
				return TagEditor.DetachModifiedTagArray();
			}

			return nullptr;
		}

		mapi::CopyDetails GetCopyDetails(
			HWND hWnd,
			_In_ LPMAPIPROP lpSource,
			LPCGUID lpGUID,
			_In_opt_ LPSPropTagArray lpTagArray,
			bool bIsAB)
		{
			dialog::editor::CEditor MyData(
				nullptr, IDS_COPYTO, IDS_COPYPASTEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.AddPane(
				viewpane::TextPane::CreateSingleLinePane(0, IDS_INTERFACE, guid::GUIDToStringAndName(lpGUID), false));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_FLAGS, false));
			MyData.SetHex(1, MAPI_DIALOG);

			if (!MyData.DisplayDialog()) return {};

			auto progress = hWnd ? mapi::mapiui::GetMAPIProgress(L"CopyTo", hWnd) : LPMAPIPROGRESS{};

			return {true,
					MyData.GetHex(1) | (progress ? MAPI_DIALOG : 0),
					guid::StringToGUID(MyData.GetStringW(0)),
					progress,
					progress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
					GetExcludedTags(lpTagArray, lpSource, bIsAB),
					true};
		}

		void init()
		{
			mapi::GetCopyDetails = [](auto _1, auto _2, auto _3, auto _4, auto _5) -> mapi::CopyDetails {
				return ui::callbacks::GetCopyDetails(_1, _2, _3, _4, _5);
			};
		}
	} // namespace callbacks
} // namespace ui