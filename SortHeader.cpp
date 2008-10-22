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
			if (m_hwndTip)
			{
				HDHITTESTINFO hdHitTestInfo = {0};

#define GET_X_LPARAM(lp)                        ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp)                        ((int)(short)HIWORD(lp))
				hdHitTestInfo.pt.x = GET_X_LPARAM(lParam);
				hdHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

				EC_B(::SendMessage(m_hWnd, HDM_HITTEST, 0, (LPARAM) &hdHitTestInfo));

				if (hdHitTestInfo.flags & HHT_ONHEADER)
				{
					HDITEM	hdItem = {0};
					hdItem.mask = HDI_LPARAM;

					EC_B(GetItem(hdHitTestInfo.iItem,&hdItem));

					LPHEADERDATA lpHeaderData = (LPHEADERDATA) hdItem.lParam;

					// this will only display tips if we have a HeaderData structure saved
					if (lpHeaderData)
					{
						m_ti.lpszText = lpHeaderData->szTipString;
						EC_B(::GetCursorPos( &hdHitTestInfo.pt ));
						EC_B(::SendMessage(m_hwndTip, TTM_TRACKPOSITION, 0, (LPARAM)MAKELPARAM(hdHitTestInfo.pt.x, hdHitTestInfo.pt.y)));
						EC_B(::SendMessage(m_hwndTip, TTM_SETTOOLINFO, true, (LPARAM) (LPTOOLINFO) &m_ti));
						EC_B(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, true, (LPARAM) (LPTOOLINFO) &m_ti));
					}
				}

				TRACKMOUSEEVENT tmEvent = {0};
				tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
				tmEvent.dwFlags = TME_LEAVE;
				tmEvent.hwndTrack = m_hWnd;

				EC_B(TrackMouseEvent(&tmEvent));
			}
			return NULL;
			break;
		}
	case WM_MOUSELEAVE:
		{
			if (m_hwndTip)
			{
				EC_B(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, false, (LPARAM) (LPTOOLINFO) &m_ti));
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
