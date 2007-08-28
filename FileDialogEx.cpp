#include "stdafx.h"
#include "FileDialogEx.h"
#include "Error.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static BOOL IsWin2000();

// To get these I should be including afxpriv but that brings in a whole
// bunch of garbage that strsafe doesn't like.
//#include <afxpriv.h>
void AFXAPI AfxHookWindowCreate(CWnd* pWnd);
BOOL AFXAPI AfxUnhookWindowCreate();

///////////////////////////////////////////////////////////////////////////
// CFileDialogEx

IMPLEMENT_DYNAMIC(CFileDialogEx, CFileDialog)

#define CCHBIGBUFF 8192
CFileDialogEx::CFileDialogEx(BOOL bOpenFileDialog,
							 LPCTSTR lpszDefExt,
							 LPCTSTR lpszFileName,
							 DWORD dwFlags, LPCTSTR lpszFilter, CWnd* pParentWnd) :
CFileDialog(bOpenFileDialog, lpszDefExt, lpszFileName,
			dwFlags, lpszFilter, pParentWnd)
{
	m_szbigBuff = NULL;
		// Modify OPENFILENAME members directly to point to m_szbigBuff
	if (dwFlags & OFN_ALLOWMULTISELECT)
	{
		m_szbigBuff = new TCHAR[CCHBIGBUFF];
		if (m_szbigBuff)
		{
			m_szbigBuff[0] = 0;// NULL terminate
			m_ofn.lpstrFile = m_szbigBuff;
			m_ofn.nMaxFile = CCHBIGBUFF;
		}
	}
}

CFileDialogEx::~CFileDialogEx()
{
	delete[] m_szbigBuff;
}

BEGIN_MESSAGE_MAP(CFileDialogEx, CFileDialog)
//{{AFX_MSG_MAP(CFileDialogEx)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BOOL IsWin2000()
{
	OSVERSIONINFOEX osvi;
	HRESULT hRes = S_OK;

	// Try calling GetVersionEx using the OSVERSIONINFOEX structure,
	// which is supported on Windows 2000.
	//
	// If that fails, try using the OSVERSIONINFO structure.

	ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));
	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);

	WC_B(GetVersionEx ((OSVERSIONINFO*) &osvi));
	if (S_OK != hRes)
	{
		// If OSVERSIONINFOEX doesn't work, try OSVERSIONINFO.
		hRes = S_OK;
		osvi.dwOSVersionInfoSize = sizeof (OSVERSIONINFO);
		EC_B(GetVersionEx((OSVERSIONINFO*) &osvi));
		if (S_OK != hRes) return FALSE;
	}

	switch (osvi.dwPlatformId)
	{
	case VER_PLATFORM_WIN32_NT:

		if (osvi.dwMajorVersion >= 5)
			return TRUE;

		break;
	}
	return FALSE;
}

INT_PTR CFileDialogEx::DoModal()
{
	ASSERT_VALID(this);
	ASSERT(m_ofn.Flags & OFN_ENABLEHOOK);
	ASSERT(m_ofn.lpfnHook != NULL); // can still be a user hook
	HRESULT hRes = S_OK;

	// zero out the file buffer for consistent parsing later
	ASSERT(AfxIsValidAddress(m_ofn.lpstrFile, m_ofn.nMaxFile));
	size_t nOffset = 0;
	EC_H(StringCchLength(m_ofn.lpstrFile,STRSAFE_MAX_CCH,&nOffset));
	nOffset++;
	ASSERT(nOffset <= m_ofn.nMaxFile);
	memset(m_ofn.lpstrFile+nOffset, 0,
		(m_ofn.nMaxFile-nOffset)*sizeof(TCHAR));

	// WINBUG: This is a special case for the file open/save dialog,
	//  which sometimes pumps while it is coming up but before it has
	//  disabled the main window.
	HWND hWndFocus = NULL;
	EC_D(hWndFocus,::GetFocus());
	BOOL bEnableParent = FALSE;
	m_ofn.hwndOwner = PreModal();
	AfxUnhookWindowCreate();
	if (m_ofn.hwndOwner != NULL && ::IsWindowEnabled(m_ofn.hwndOwner))
	{
		bEnableParent = TRUE;
		::EnableWindow(m_ofn.hwndOwner, FALSE);//Don't wrap in EC_B
	}

	_AFX_THREAD_STATE* pThreadState = AfxGetThreadState();
	ASSERT(pThreadState->m_pAlternateWndInit == NULL);

	if (m_ofn.Flags & OFN_EXPLORER)
		pThreadState->m_pAlternateWndInit = this;
	else
		AfxHookWindowCreate(this);

	memset(&m_ofnEx, 0, sizeof(OPENFILENAMEEX));

	// m_ofn is always the NT4 sized struct
	// m_ofnEx is always the win2k sized struct
	// use OPENFILENAME_SIZE_VERSION_400 to ensure we only copy the small struct
	memcpy(&m_ofnEx, &m_ofn, OPENFILENAME_SIZE_VERSION_400);
	if (IsWin2000())
		m_ofnEx.lStructSize = sizeof(OPENFILENAMEEX);

	int nResult;
	if (m_bOpenFileDialog)
		nResult = ::GetOpenFileName((OPENFILENAME*)&m_ofnEx);
	else
		nResult = ::GetSaveFileName((OPENFILENAME*)&m_ofnEx);

	memcpy(&m_ofn, &m_ofnEx, OPENFILENAME_SIZE_VERSION_400);
	m_ofn.lStructSize = OPENFILENAME_SIZE_VERSION_400;

	if (nResult)
		ASSERT(pThreadState->m_pAlternateWndInit == NULL);
	pThreadState->m_pAlternateWndInit = NULL;

	// WINBUG: Second part of special case for file open/save dialog.
	if (bEnableParent)
		::EnableWindow(m_ofnEx.hwndOwner, TRUE);//Don't wrap in EC_B
	if (::IsWindow(hWndFocus))
		::SetFocus(hWndFocus);

	PostModal();
	return nResult ? nResult : IDCANCEL;
}

BOOL CFileDialogEx::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	memcpy(&m_ofn, &m_ofnEx, OPENFILENAME_SIZE_VERSION_400);
	m_ofn.lStructSize = OPENFILENAME_SIZE_VERSION_400;

	return CFileDialog::OnNotify(wParam, lParam, pResult);
}
