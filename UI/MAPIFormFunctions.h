#pragma once
// Stand alone MAPI Form functions

namespace controls
{
	namespace sortlistctrl
	{
		class CContentsTableListCtrl;
	}
} // namespace controls

namespace mapi
{
	namespace mapiui
	{
		_Check_return_ HRESULT CreateAndDisplayNewMailInFolder(
			_In_ HWND hwndParent,
			_In_ LPMDB lpMDB,
			_In_ LPMAPISESSION lpMAPISession,
			_In_ controls::sortlistctrl::CContentsTableListCtrl* lpContentsTableListCtrl,
			int iItem,
			_In_ const std::wstring& szMessageClass,
			_In_ LPMAPIFOLDER lpFolder);
		_Check_return_ HRESULT OpenMessageModal(
			_In_ LPMAPIFOLDER lpParentFolder,
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMDB lpMDB,
			_In_ LPMESSAGE lpMessage);
		_Check_return_ HRESULT OpenMessageNonModal(
			_In_ HWND hwndParent,
			_In_ LPMDB lpMDB,
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMAPIFOLDER lpSourceFolder,
			_In_ controls::sortlistctrl::CContentsTableListCtrl* lpContentsTableListCtrl,
			int iItem,
			_In_ LPMESSAGE lpMessage,
			LONG lVerb,
			_In_opt_ LPCRECT lpRect);
	} // namespace mapiui
} // namespace mapi