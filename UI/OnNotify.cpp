#include <StdAfx.h>
#include <UI/OnNotify.h>
#include <core/utility/error.h>

namespace mapi::mapiui
{
	void OnNotify(HWND hWndParent, HTREEITEM hTreeParent, ULONG cNotify, LPNOTIFICATION lpNotifications)
	{
		if (!hWndParent) return;

		for (ULONG i = 0; i < cNotify; i++)
		{
			if (fnevNewMail == lpNotifications[i].ulEventType)
			{
				MessageBeep(MB_OK);
			}
			else if (fnevTableModified == lpNotifications[i].ulEventType)
			{
				switch (lpNotifications[i].info.tab.ulTableEvent)
				{
				case TABLE_ERROR:
				case TABLE_CHANGED:
				case TABLE_RELOAD:
					EC_H_S(static_cast<HRESULT>(
						::SendMessage(hWndParent, WM_MFCMAPI_REFRESHTABLE, reinterpret_cast<WPARAM>(hTreeParent), 0)));
					break;
				case TABLE_ROW_ADDED:
					EC_H_S(static_cast<HRESULT>(::SendMessage(
						hWndParent,
						WM_MFCMAPI_ADDITEM,
						reinterpret_cast<WPARAM>(&lpNotifications[i].info.tab),
						reinterpret_cast<LPARAM>(hTreeParent))));
					break;
				case TABLE_ROW_DELETED:
					EC_H_S(static_cast<HRESULT>(::SendMessage(
						hWndParent,
						WM_MFCMAPI_DELETEITEM,
						reinterpret_cast<WPARAM>(&lpNotifications[i].info.tab),
						reinterpret_cast<LPARAM>(hTreeParent))));
					break;
				case TABLE_ROW_MODIFIED:
					EC_H_S(static_cast<HRESULT>(::SendMessage(
						hWndParent,
						WM_MFCMAPI_MODIFYITEM,
						reinterpret_cast<WPARAM>(&lpNotifications[i].info.tab),
						reinterpret_cast<LPARAM>(hTreeParent))));
					break;
				}
			}
		}
	}
} // namespace mapi::mapiui