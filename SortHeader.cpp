// SortHeader.cpp : implementation file
//

#include "stdafx.h"
#include "SortHeader.h"
#include "InterpretProp2.h"

CSortHeader::CSortHeader()
{
	m_hwndTip = NULL;
	m_hwndParent = NULL;
	ZeroMemory(&m_ti,sizeof(TOOLINFO));
}

BEGIN_MESSAGE_MAP(CSortHeader, CHeaderCtrl)
	ON_MESSAGE(WM_MFCMAPI_SAVECOLUMNORDERHEADER, msgOnSaveColumnOrder)
END_MESSAGE_MAP()

BOOL CSortHeader::Init(CHeaderCtrl *pHeader, HWND hwndParent)
{
	if (!pHeader) return false;

	if (!SubclassWindow(pHeader->GetSafeHwnd()))
	{
		DebugPrint(DBGCreateDialog,_T("CSortHeader::Init Unable to subclass existing header!\n"));
		return false;
	}
	m_hwndParent = hwndParent;
	m_bTrackingMouse = false;

	RegisterHeaderTooltip();

	return true;
}

void CSortHeader::RegisterHeaderTooltip()
{
	if (!m_hwndTip)
	{
		HRESULT hRes = S_OK;
		EC_D(m_hwndTip,CreateWindowEx(
			WS_EX_TOPMOST,
			TOOLTIPS_CLASS,
			NULL,
			TTS_NOPREFIX | TTS_ALWAYSTIP | TTS_BALLOON,
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
}

LRESULT CSortHeader::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	HRESULT hRes = S_OK;

	switch (message)
	{
	case WM_MOUSEMOVE:
		{
			if (m_hwndTip)
			{
				static int oldX = 0;
				static int oldY = 0;
#define GET_X_LPARAM(lp)                        ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp)                        ((int)(short)HIWORD(lp))
				int newX = GET_X_LPARAM(lParam);
				int newY = GET_Y_LPARAM(lParam);

				HDHITTESTINFO hdHitTestInfo = {0};

				hdHitTestInfo.pt.x = newX;
				hdHitTestInfo.pt.y = newY;

				EC_B(::SendMessage(m_hWnd, HDM_HITTEST, 0, (LPARAM) &hdHitTestInfo));

				// We only turn on or modify our tooltip if we're on a column header
				if (hdHitTestInfo.flags & HHT_ONHEADER)
				{
					// Make sure the mouse has actually moved. The presence of the ToolTip
					// causes Windows to send the message continuously.
					if ((newX != oldX) || (newY != oldY))
					{
						oldX = newX;
						oldY = newY;

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
							if (!m_bTrackingMouse)
							{
								// Ask for notification when the mouse leaves the control
								TRACKMOUSEEVENT tmEvent = {0};
								tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
								tmEvent.dwFlags = TME_LEAVE;
								tmEvent.hwndTrack = m_hWnd;
								EC_B(TrackMouseEvent(&tmEvent));

								// Turn on the tooltip
								EC_B(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, true, (LPARAM) (LPTOOLINFO) &m_ti));
								m_bTrackingMouse = true;
							}
						}
					}
				}
				else if (m_bTrackingMouse)
				{
					// We were displaying a tooltip, but we're now on empty space in the header. Turn it off.
					EC_B(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, false, (LPARAM) (LPTOOLINFO) &m_ti));
					m_bTrackingMouse = false;
				}
			}
			break;
		}
	case WM_MOUSELEAVE:
		{
			if (m_hwndTip)
			{
				// We were displaying a tooltip, but we're now off of the header. Turn it off.
				EC_B(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, false, (LPARAM) (LPTOOLINFO) &m_ti));
				m_bTrackingMouse = false;
			}
			return NULL;
			break;
		}
	} // end switch
	return CHeaderCtrl::WindowProc(message,wParam,lParam);
}

// WM_MFCMAPI_SAVECOLUMNORDERHEADER
LRESULT	CSortHeader::msgOnSaveColumnOrder(WPARAM /*wParam*/, LPARAM /*lParam*/)
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
}
