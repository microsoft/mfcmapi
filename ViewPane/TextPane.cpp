#include "stdafx.h"
#include "..\stdafx.h"
#include "TextPane.h"
#include "..\MAPIFunctions.h"
#include "..\String.h"
#include "..\UIFunctions.h"

static wstring CLASS = L"TextPane";

ViewPane* CreateMultiLinePaneA(UINT uidLabel, _In_opt_z_ LPCSTR szVal, bool bReadOnly)
{
	return CreateSingleLinePaneA(uidLabel, szVal, bReadOnly, true);
}

ViewPane* CreateMultiLinePaneW(UINT uidLabel, _In_opt_z_ LPCWSTR szVal, bool bReadOnly)
{
	return CreateSingleLinePaneW(uidLabel, szVal, bReadOnly, true);
}

ViewPane* CreateSingleLinePaneA(UINT uidLabel, _In_opt_z_ LPCSTR szVal, bool bReadOnly, bool bMultiLine)
{
	TextPane* lpPane = new TextPane(uidLabel, bReadOnly, bMultiLine);
	if (lpPane)
	{
		lpPane->SetStringA(szVal);
	}
	return lpPane;
}

ViewPane* CreateSingleLinePaneW(UINT uidLabel, _In_opt_z_ LPCWSTR szVal, bool bReadOnly, bool bMultiLine)
{
	TextPane* lpPane = new TextPane(uidLabel, bReadOnly, bMultiLine);
	if (lpPane)
	{
		lpPane->SetStringW(szVal);
	}
	return lpPane;
}

ViewPane* CreateSingleLinePaneID(UINT uidLabel, UINT uidVal, bool bReadOnly)
{
	TextPane* lpPane = new TextPane(uidLabel, bReadOnly, false);

	if (lpPane)
	{
		wstring szTemp;

		if (uidVal)
		{
			szTemp = loadstring(uidVal);
		}

		lpPane->SetStringW(szTemp.c_str());
	}

	return lpPane;
}

// Imports binary data from a stream, converting it to hex format before returning
_Check_return_ static DWORD CALLBACK EditStreamReadCallBack(
	DWORD_PTR dwCookie,
	_In_ LPBYTE pbBuff,
	LONG cb,
	_In_count_(cb) LONG *pcb)
{
	HRESULT hRes = S_OK;
	if (!pbBuff || !pcb || !dwCookie) return 0;

	LPSTREAM stmData = (LPSTREAM)dwCookie;

	*pcb = 0;

	DebugPrint(DBGStream, L"EditStreamReadCallBack: cb = %d\n", cb);

	LONG cbTemp = cb / 2;
	ULONG cbTempRead = 0;
	LPBYTE pbTempBuff = new BYTE[cbTemp];

	if (pbTempBuff)
	{
		EC_MAPI(stmData->Read(pbTempBuff, cbTemp, &cbTempRead));
		DebugPrint(DBGStream, L"EditStreamReadCallBack: read %u bytes\n", cbTempRead);

		memset(pbBuff, 0, cbTempRead * 2);
		ULONG i = 0;
		ULONG iBinPos = 0;
		for (i = 0; i < cbTempRead; i++)
		{
			BYTE bLow;
			BYTE bHigh;
			CHAR szLow;
			CHAR szHigh;

			bLow = (BYTE)((pbTempBuff[i]) & 0xf);
			bHigh = (BYTE)((pbTempBuff[i] >> 4) & 0xf);
			szLow = (CHAR)((bLow <= 0x9) ? '0' + bLow : 'A' + bLow - 0xa);
			szHigh = (CHAR)((bHigh <= 0x9) ? '0' + bHigh : 'A' + bHigh - 0xa);

			pbBuff[iBinPos] = szHigh;
			pbBuff[iBinPos + 1] = szLow;

			iBinPos += 2;
		}

		*pcb = cbTempRead * 2;

		delete[] pbTempBuff;
	}

	return 0;
} // EditStreamReadCallBack

TextPane::TextPane(UINT uidLabel, bool bReadOnly, bool bMultiLine) :ViewPane(uidLabel, bReadOnly)
{
	m_lpszW = NULL;
	m_lpszA = NULL;
	m_cchsz = 0;
	m_bMultiline = bMultiLine;
}

TextPane::~TextPane()
{
	ClearString();
}

bool TextPane::IsType(__ViewTypes vType)
{
	return CTRL_TEXTPANE == vType;
}

ULONG TextPane::GetFlags()
{
	ULONG ulFlags = vpNone;
	if (m_EditBox.m_hWnd && m_EditBox.GetModify()) ulFlags |= vpDirty;
	if (m_bReadOnly) ulFlags |= vpReadonly;
	return ulFlags;
}

int TextPane::GetFixedHeight()
{
	int iHeight = 0;
	if (0 != m_iControl) iHeight += m_iSmallHeightMargin; // Top margin

	if (m_bUseLabelControl)
	{
		// Text labels will bump directly against their edit control, so we don't add a margin here
		iHeight += m_iLabelHeight;
	}

	if (!m_bMultiline)
	{
		iHeight += m_iEditHeight;
	}

	iHeight += m_iSmallHeightMargin; // Bottom margin

	return iHeight;
}

int TextPane::GetLines()
{
	if (m_bMultiline)
	{
		return LINES_MULTILINEEDIT;
	}
	else
	{
		return 0;
	}
}

void TextPane::SetWindowPos(int x, int y, int width, int height)
{
	HRESULT hRes = S_OK;
	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
		height -= m_iSmallHeightMargin;
	}

	if (m_bUseLabelControl)
	{
		EC_B(m_Label.SetWindowPos(
			0,
			x,
			y,
			width,
			m_iLabelHeight,
			SWP_NOZORDER));
		y += m_iLabelHeight;
		height -= m_iLabelHeight;
	}

	height -= m_iSmallHeightMargin; // This is the bottom margin

	EC_B(m_EditBox.SetWindowPos(NULL, x, y, width, height, SWP_NOZORDER));
}

void TextPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC /*hdc*/)
{
	ViewPane::Initialize(iControl, pParent, NULL);

	HRESULT hRes = S_OK;

	EC_B(m_EditBox.Create(
		WS_TABSTOP
		| WS_CHILD
		| WS_CLIPSIBLINGS
		| WS_BORDER
		| WS_VISIBLE
		| WS_VSCROLL
		| ES_AUTOVSCROLL
		| (m_bReadOnly ? ES_READONLY : 0)
		| (m_bMultiline ? (ES_MULTILINE | ES_WANTRETURN) : (ES_AUTOHSCROLL)),
		CRect(0, 0, 0, 0),
		pParent,
		m_nID));
	SubclassEdit(m_EditBox.m_hWnd, pParent ? pParent->m_hWnd : NULL, m_bReadOnly);

	m_bInitialized = true; // We can now call SetEditBoxText

	// Set maximum text size
	// Use -1 to allow for VERY LARGE strings
	(void) ::SendMessage(
		m_EditBox.m_hWnd,
		EM_EXLIMITTEXT,
		(WPARAM)0,
		(LPARAM)-1);

	SetEditBoxText();

	m_EditBox.SetEventMask(ENM_CHANGE);

	// Clear the modify bits so we can detect changes later
	m_EditBox.SetModify(false);

	// Remove the awful autoselect of the edit control that scrolls to the end of multiline text
	if (m_bMultiline)
	{
		::PostMessage(
			m_EditBox.m_hWnd,
			EM_SETSEL,
			(WPARAM)0,
			(LPARAM)0);
	}
}

struct FakeStream
{
	LPWSTR lpszW;
	size_t cbszW;
	size_t cbCur;
};

_Check_return_ static DWORD CALLBACK FakeEditStreamReadCallBack(
	DWORD_PTR dwCookie,
	_In_ LPBYTE pbBuff,
	LONG cb,
	_In_count_(cb) LONG *pcb)
{
	if (!pbBuff || !pcb || !dwCookie) return 0;

	FakeStream* lpfs = (FakeStream*)dwCookie;
	if (!lpfs) return 0;
	ULONG cbRemaining = (ULONG)(lpfs->cbszW - lpfs->cbCur);
	ULONG cbRead = min((ULONG)cb, cbRemaining);

	*pcb = cbRead;

	if (cbRead) memcpy(pbBuff, ((LPBYTE)lpfs->lpszW) + lpfs->cbCur, cbRead);

	lpfs->cbCur += cbRead;

	return 0;
}

void TextPane::SetEditBoxText()
{
	if (!m_bInitialized) return;
	if (!m_EditBox.m_hWnd) return;

	ULONG ulEventMask = m_EditBox.GetEventMask(); // Get original mask
	m_EditBox.SetEventMask(ENM_NONE);

	// In order to support strings with embedded NULLs, we're going to stream the string in
	// We don't have to build a real stream interface - we can fake a lightweight one
	EDITSTREAM es = { 0, 0, FakeEditStreamReadCallBack };
	UINT uFormat = SF_TEXT | SF_UNICODE;

	FakeStream fs = { 0 };
	fs.lpszW = m_lpszW;
	fs.cbszW = m_cchsz * sizeof(WCHAR);

	// The edit control's gonna read in the actual NULL terminator, which we do not want, so back off one character
	if (fs.cbszW) fs.cbszW -= sizeof(WCHAR);

	es.dwCookie = (DWORD_PTR)&fs;

	// read the 'text stream' into control
	long lBytesRead = 0;
	lBytesRead = m_EditBox.StreamIn(uFormat, es);
	DebugPrintEx(DBGStream, CLASS, L"SetEditBoxText", L"read %d bytes from the stream\n", lBytesRead);

	// Clear the modify bit so this stream appears untouched
	m_EditBox.SetModify(false);

	m_EditBox.SetEventMask(ulEventMask); // put original mask back
}

// Clears the strings out of an lpEdit
void TextPane::ClearString()
{
	delete[] m_lpszW;
	delete[] m_lpszA;
	m_lpszW = NULL;
	m_lpszA = NULL;
	m_cchsz = NULL;
}

// Sets m_lpControls[i].UI.lpEdit->lpszW using SetStringW
// cchsz of -1 lets AnsiToUnicode and SetStringW calculate the length on their own
void TextPane::SetStringA(_In_opt_z_ LPCSTR szMsg, size_t cchsz)
{
	if (!szMsg) szMsg = "";
	HRESULT hRes = S_OK;

	LPWSTR szMsgW = NULL;
	size_t cchszW = 0;
	EC_H(AnsiToUnicode(szMsg, &szMsgW, &cchszW, cchsz));
	if (SUCCEEDED(hRes))
	{
		SetStringW(szMsgW, cchszW);
	}

	delete[] szMsgW;
}

// Sets m_lpControls[i].UI.lpEdit->lpszW
// cchsz may or may not include the NULL terminator (if one is present)
// If it is missing, we'll make sure we add it
void TextPane::SetStringW(_In_opt_z_ LPCWSTR szMsg, size_t cchsz)
{
	ClearString();

	if (!szMsg) szMsg = L"";
	HRESULT hRes = S_OK;
	size_t cchszW = cchsz;

	if (-1 == cchszW)
	{
		EC_H(StringCchLengthW(szMsg, STRSAFE_MAX_CCH, &cchszW));
	}
	// If we were passed a NULL terminated string,
	// cchszW counts the NULL terminator. Back off one.
	else if (cchszW && szMsg[cchszW - 1] == NULL)
	{
		cchszW--;
	}
	// cchszW is now the length of our string not counting the NULL terminator

	// add one for a NULL terminator
	m_cchsz = cchszW + 1;
	m_lpszW = new WCHAR[cchszW + 1];

	if (m_lpszW)
	{
		memcpy(m_lpszW, szMsg, cchszW * sizeof(WCHAR));
		m_lpszW[cchszW] = NULL;
	}

	SetEditBoxText();
}

void TextPane::SetBinary(_In_opt_count_(cb) LPBYTE lpb, size_t cb)
{
	if (!lpb || !cb)
	{
		SetStringW(NULL);
	}
	else
	{
		SetStringW(BinToHexString(lpb, cb, false).c_str());
	}
}

// This is used by the DbgView - don't call any debugger functions here!!!
void TextPane::AppendString(_In_z_ LPCTSTR szMsg)
{
	m_EditBox.HideSelection(false, true);
	GETTEXTLENGTHEX getTextLength = { 0 };
	getTextLength.flags = GTL_PRECISE | GTL_NUMCHARS;
	getTextLength.codepage = 1200;

	int cchText = m_EditBox.GetWindowTextLength();
	m_EditBox.SetSel(cchText, cchText);
	m_EditBox.ReplaceSel(szMsg);
}

// This is used by the DbgView - don't call any debugger functions here!!!
void TextPane::ClearView()
{
	ClearString();
	::SendMessage(
		m_EditBox.m_hWnd,
		WM_SETTEXT,
		NULL,
		(LPARAM)_T(""));
}

void TextPane::SetEditReadOnly()
{
	m_EditBox.SetBackgroundColor(false, MyGetSysColor(cBackgroundReadOnly));
	m_EditBox.SetReadOnly();
}

LPWSTR TextPane::GetStringW()
{
	return m_lpszW;
}

LPSTR TextPane::GetStringA()
{
	HRESULT hRes = S_OK;
	// Don't use ClearString - that would wipe the lpszW too!
	delete[] m_lpszA;
	m_lpszA = NULL;

	// We're not leaking this conversion
	// It goes into m_lpControls[i].UI.lpEdit->lpszA, which we manage
	EC_H(UnicodeToAnsi(m_lpszW, &m_lpszA, m_cchsz));

	return m_lpszA;
}

// Gets string from edit box and places it in m_lpControls[i].UI.lpEdit->lpszW
void TextPane::CommitUIValues()
{
	ClearString();

	GETTEXTLENGTHEX getTextLength = { 0 };
	getTextLength.flags = GTL_PRECISE | GTL_NUMCHARS;
	getTextLength.codepage = 1200;

	size_t cchText = 0;

	cchText = (size_t)::SendMessage(
		m_EditBox.m_hWnd,
		EM_GETTEXTLENGTHEX,
		(WPARAM)&getTextLength,
		(LPARAM)0);
	if (E_INVALIDARG == cchText)
	{
		// we didn't get a length - try another method
		cchText = (size_t)::SendMessage(
			m_EditBox.m_hWnd,
			WM_GETTEXTLENGTH,
			(WPARAM)0,
			(LPARAM)0);
	}

	// cchText will never include the NULL terminator, so add one to our count
	cchText += 1;

	LPWSTR lpszW = (WCHAR*) new WCHAR[cchText];
	size_t cchW = 0;

	if (lpszW)
	{
		memset(lpszW, 0, cchText * sizeof(WCHAR));

		if (cchText > 1) // No point in checking if the string is just a null terminator
		{
			GETTEXTEX getText = { 0 };
			getText.cb = (DWORD)cchText * sizeof(WCHAR);
			getText.flags = GT_DEFAULT;
			getText.codepage = 1200;

			cchW = ::SendMessage(
				m_EditBox.m_hWnd,
				EM_GETTEXTEX,
				(WPARAM)&getText,
				(LPARAM)lpszW);
			if (0 == cchW)
			{
				// Didn't get a string from this message, fall back to WM_GETTEXT
				LPSTR lpszA = (CHAR*) new CHAR[cchText];
				if (lpszA)
				{
					memset(lpszA, 0, cchText * sizeof(CHAR));
					HRESULT hRes = S_OK;
					cchW = ::SendMessage(
						m_EditBox.m_hWnd,
						WM_GETTEXT,
						(WPARAM)cchText,
						(LPARAM)lpszA);
					if (0 != cchW)
					{
						EC_H(StringCchPrintfW(lpszW, cchText, L"%hs", lpszA)); // STRING_OK
					}
					delete[] lpszA;
				}
			}
		}

		m_lpszW = lpszW;
		m_cchsz = cchText;
	}
}

// No need to free this - treat it like a static
_Check_return_ LPSTR TextPane::GetEditBoxTextA(_Out_ size_t* lpcchText)
{
	CommitUIValues();
	if (lpcchText) *lpcchText = m_cchsz;
	return GetStringA();
}

_Check_return_ LPWSTR TextPane::GetEditBoxTextW(_Out_ size_t* lpcchText)
{
	CommitUIValues();
	if (lpcchText) *lpcchText = m_cchsz;
	return GetStringW();
}

_Check_return_ wstring TextPane::GetStringUseControl()
{
	int len = GetWindowTextLength(m_EditBox.m_hWnd) + 1;
	vector<wchar_t> text(len);
	GetWindowTextW(m_EditBox.m_hWnd, &text[0], len);
	return &text[0];
}

// Takes a binary stream and initializes an edit control with the HEX version of this stream
void TextPane::InitEditFromBinaryStream(_In_ LPSTREAM lpStreamIn)
{
	EDITSTREAM es = { 0, 0, EditStreamReadCallBack };
	UINT uFormat = SF_TEXT;

	long lBytesRead = 0;

	es.dwCookie = (DWORD_PTR)lpStreamIn;

	// read the 'text' stream into control
	lBytesRead = m_EditBox.StreamIn(uFormat, es);
	DebugPrintEx(DBGStream, CLASS, L"InitEditFromStream", L"read %d bytes from the stream\n", lBytesRead);

	// Clear the modify bit so this stream appears untouched
	m_EditBox.SetModify(false);
}

// Writes a hex pane out to a binary stream
void TextPane::WriteToBinaryStream(_In_ LPSTREAM lpStreamOut)
{
	HRESULT hRes = S_OK;

	auto bin = HexStringToBin(GetStringUseControl());
	if (bin.data() != 0)
	{
		ULONG cbWritten = 0;
		EC_MAPI(lpStreamOut->Write(bin.data(), (ULONG)bin.size(), &cbWritten));
		DebugPrintEx(DBGStream, CLASS, L"WriteToBinaryStream", L"wrote 0x%X bytes to the stream\n", cbWritten);

		EC_MAPI(lpStreamOut->Commit(STGC_DEFAULT));
	}
}

void TextPane::ShowWindow(int nCmdShow)
{
	m_EditBox.ShowWindow(nCmdShow);
}
