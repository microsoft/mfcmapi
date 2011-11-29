// SortHeader.cpp : implementation file
//

#include "stdafx.h"
#include "SortHeader.h"
#include "UIFunctions.h"

CSortHeader::CSortHeader()
{
	m_hwndTip = NULL;
	m_hwndParent = NULL;
	m_bInTrack = false;
	m_iTrack = -1;
	m_iHeaderHeight = 0;
	ZeroMemory(&m_ti,sizeof(TOOLINFO));
} // CSortHeader::CSortHeader

BEGIN_MESSAGE_MAP(CSortHeader, CHeaderCtrl)
	ON_MESSAGE(WM_MFCMAPI_SAVECOLUMNORDERHEADER, msgOnSaveColumnOrder)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, OnCustomDraw)
	ON_NOTIFY_REFLECT(HDN_BEGINTRACK, OnBeginTrack)
	ON_NOTIFY_REFLECT(HDN_ENDTRACK, OnEndTrack)
	ON_NOTIFY_REFLECT(HDN_ITEMCHANGED, OnTrack)
END_MESSAGE_MAP()

_Check_return_ bool CSortHeader::Init(_In_ CHeaderCtrl *pHeader, _In_ HWND hwndParent)
{
	if (!pHeader) return false;

	if (!SubclassWindow(pHeader->GetSafeHwnd()))
	{
		DebugPrint(DBGCreateDialog,_T("CSortHeader::Init Unable to subclass existing header!\n"));
		return false;
	}
	m_hwndParent = hwndParent;
	m_bTooltipDisplayed = false;

	RegisterHeaderTooltip();

	return true;
} // CSortHeader::Init

void CSortHeader::RegisterHeaderTooltip()
{
	if (!m_hwndTip)
	{
		HRESULT hRes = S_OK;
		EC_D(m_hwndTip,CreateWindowEx(
			WS_EX_TOPMOST,
			TOOLTIPS_CLASS,
			NULL,
			TTS_NOPREFIX | TTS_ALWAYSTIP,
			CW_USEDEFAULT,
			CW_USEDEFAULT,
			CW_USEDEFAULT,
			CW_USEDEFAULT,
			m_hWnd,
			NULL,
			AfxGetInstanceHandle(),
			NULL));

		if (m_hwndTip)
		{
			EC_B(::SetWindowPos(
				m_hwndTip,
				HWND_TOPMOST,
				0, 0, 0, 0,
				SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE));

			m_ti.cbSize = sizeof(TOOLINFO);
			m_ti.uFlags = TTF_TRACK | TTF_IDISHWND;
			m_ti.hwnd = m_hWnd;
			m_ti.uId = (UINT_PTR) m_hWnd;
			m_ti.hinst = AfxGetInstanceHandle();
			m_ti.lpszText = _T("");

			EC_B(::SendMessage(m_hwndTip, TTM_ADDTOOL, 0, (LPARAM) (LPTOOLINFO) &m_ti));
			EC_B(::SendMessage(m_hwndTip, TTM_SETMAXTIPWIDTH, 0, (LPARAM) 500));
		}
	}
} // CSortHeader::RegisterHeaderTooltip

#define GET_X_LPARAM(lp) ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp) ((int)(short)HIWORD(lp))
_Check_return_ LRESULT CSortHeader::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	HRESULT hRes = S_OK;

	if (m_hwndTip)
	{
		switch (message)
		{
		case WM_ERASEBKGND:
			{
				return true;
				break;
			}
		case WM_MOUSEMOVE:
			{
				// This covers the case where we move from one header to another, but don't leave the control
				if (m_bTooltipDisplayed)
				{
					HDHITTESTINFO hdHitTestInfo = {0};
					hdHitTestInfo.pt.x = GET_X_LPARAM(lParam);
					hdHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

					WC_B(::SendMessage(m_hWnd, HDM_HITTEST, 0, (LPARAM) &hdHitTestInfo));
					if (!(hdHitTestInfo.flags & HHT_ONHEADER))
					{
						// We were displaying a tooltip, but we're now on empty space in the header. Turn it off.
						EC_B(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, false, (LPARAM) (LPTOOLINFO) &m_ti));
						m_bTooltipDisplayed = false;
					}
				}
				else
				{
					TRACKMOUSEEVENT tmEvent = {0};
					tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
					tmEvent.dwFlags = TME_HOVER;
					tmEvent.hwndTrack = m_hWnd;
					tmEvent.dwHoverTime = HOVER_DEFAULT;

					EC_B(TrackMouseEvent(&tmEvent));
				}
				break;
			}
		case WM_MOUSEHOVER:
			{
				HDHITTESTINFO hdHitTestInfo = {0};
				hdHitTestInfo.pt.x = GET_X_LPARAM(lParam);
				hdHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

				EC_B(::SendMessage(m_hWnd, HDM_HITTEST, 0, (LPARAM) &hdHitTestInfo));

				// We only turn on or modify our tooltip if we're on a column header
				if (hdHitTestInfo.flags & HHT_ONHEADER)
				{
					HDITEM	hdItem = {0};
					hdItem.mask = HDI_LPARAM;

					EC_B(GetItem(hdHitTestInfo.iItem,&hdItem));

					LPHEADERDATA lpHeaderData = (LPHEADERDATA) hdItem.lParam;

					// This will only display tips if we have a HeaderData structure saved
					if (lpHeaderData)
					{
						m_ti.lpszText = lpHeaderData->szTipString;
						EC_B(::GetCursorPos( &hdHitTestInfo.pt ));
						EC_B(::SendMessage(m_hwndTip, TTM_TRACKPOSITION, 0, (LPARAM)MAKELPARAM(hdHitTestInfo.pt.x+10, hdHitTestInfo.pt.y+20)));
						EC_B(::SendMessage(m_hwndTip, TTM_SETTOOLINFO, true, (LPARAM) (LPTOOLINFO) &m_ti));
						// Ask for notification when the mouse leaves the control
						TRACKMOUSEEVENT tmEvent = {0};
						tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
						tmEvent.dwFlags = TME_LEAVE;
						tmEvent.hwndTrack = m_hWnd;
						EC_B(TrackMouseEvent(&tmEvent));

						// Turn on the tooltip
						EC_B(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, true, (LPARAM) (LPTOOLINFO) &m_ti));
						m_bTooltipDisplayed = true;
					}
				}
				break;
			}
		case WM_MOUSELEAVE:
			{
				// We were displaying a tooltip, but we're now off of the header. Turn it off.
				EC_B(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, false, (LPARAM) (LPTOOLINFO) &m_ti));
				m_bTooltipDisplayed = false;
				return NULL;
				break;
			}
		} // end switch
	}
	return CHeaderCtrl::WindowProc(message,wParam,lParam);
} // CSortHeader::WindowProc

// WM_MFCMAPI_SAVECOLUMNORDERHEADER
_Check_return_ LRESULT CSortHeader::msgOnSaveColumnOrder(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	if (m_hwndParent)
	{
		HRESULT hRes = S_OK;
		WC_B(::PostMessage(
			m_hwndParent,
			WM_MFCMAPI_SAVECOLUMNORDERLIST,
			(WPARAM) NULL,
			(LPARAM) NULL));
	}
	return S_OK;
} // CSortHeader::msgOnSaveColumnOrder

void CSortHeader::OnBeginTrack(_In_ NMHDR* pNMHDR, _In_ LRESULT* /*pResult*/)
{
	RECT rcHeader = {0};
	if (!pNMHDR) return;
	LPNMHEADER pHdr = (LPNMHEADER) pNMHDR;
	Header_GetItemRect(pHdr->hdr.hwndFrom, pHdr->iItem, &rcHeader);
	m_bInTrack = true;
	m_iTrack = rcHeader.right;
	m_iHeaderHeight = rcHeader.bottom-rcHeader.top;
	DrawTrackingBar(pHdr->hdr.hwndFrom, m_hwndParent, m_iTrack, m_iHeaderHeight, false);
} // CSortHeader::OnBeginTrack

void CSortHeader::OnEndTrack(_In_ NMHDR* pNMHDR, _In_ LRESULT* /*pResult*/)
{
	if (m_bInTrack)
	{
		DrawTrackingBar(pNMHDR->hwndFrom, m_hwndParent, m_iTrack, m_iHeaderHeight, true);
	}
	m_bInTrack = false;
} // CSortHeader::OnEndTrack

void CSortHeader::OnTrack(_In_ NMHDR* pNMHDR, _In_ LRESULT* /*pResult*/)
{
	if (m_bInTrack && pNMHDR)
	{
		RECT rcHeader = {0};
		LPNMHEADER pHdr = (LPNMHEADER) pNMHDR;
		Header_GetItemRect(pHdr->hdr.hwndFrom, pHdr->iItem, &rcHeader);
		if (m_iTrack != rcHeader.right)
		{
			DrawTrackingBar(pHdr->hdr.hwndFrom, m_hwndParent, m_iTrack, m_iHeaderHeight, true);
			m_iTrack = rcHeader.right;
			DrawTrackingBar(pHdr->hdr.hwndFrom, m_hwndParent, m_iTrack, m_iHeaderHeight, false);
		}
	}
} // CSortHeader::OnTrack

void CSortHeader::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	CustomDrawHeader(pNMHDR, pResult);
} // CSortHeader::OnCustomDraw