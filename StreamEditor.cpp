// StreamEditor.cpp : implementation file
//

#include "stdafx.h"
#include "StreamEditor.h"
#include "InterpretProp2.h"
#include "MAPIFunctions.h"
#include "ExtraPropTags.h"
#include "SmartView\SmartView.h"

enum __StreamEditorTypes
{
	EDITOR_RTF,
	EDITOR_STREAM_BINARY,
	EDITOR_STREAM_ANSI,
	EDITOR_STREAM_UNICODE,
};

static TCHAR* CLASS = _T("CStreamEditor");

ULONG PreferredStreamType(ULONG ulPropTag)
{
	ULONG ulPropType = PROP_TYPE(ulPropTag);

	if (PT_ERROR != ulPropType && PT_UNSPECIFIED != ulPropType) return ulPropTag;

	ULONG ulPropID = PROP_ID(ulPropTag);

	switch (ulPropID)
	{
	case PROP_ID(PR_BODY):
		ulPropType = PT_TSTRING;
		break;
	case PROP_ID(PR_BODY_HTML):
		ulPropType = PT_TSTRING;
		break;
	case PROP_ID(PR_RTF_COMPRESSED):
		ulPropType = PT_BINARY;
		break;
	case PROP_ID(PR_ROAMING_BINARYSTREAM):
		ulPropType = PT_BINARY;
		break;
	default:
		ulPropType = PT_BINARY;
		break;
	}

	ulPropTag = CHANGE_PROP_TYPE(ulPropTag, ulPropType);
	return ulPropTag;
}

// Create an editor for a MAPI property - can be used to initialize stream and rtf stream editing as well
// Takes LPMAPIPROP and ulPropTag as input - will pull SPropValue from the LPMAPIPROP
CStreamEditor::CStreamEditor(
	_In_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	_In_ LPMAPIPROP lpMAPIProp,
	ULONG ulPropTag,
	bool bGuessType,
	bool bIsAB,
	bool bEditPropAsRTF,
	bool bUseWrapEx,
	ULONG ulRTFFlags,
	ULONG ulInCodePage,
	ULONG ulOutCodePage) :
	CEditor(pParentWnd, uidTitle, uidPrompt, 0, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_lpMAPIProp = lpMAPIProp;
	m_ulPropTag = ulPropTag;
	m_bAllowTypeGuessing = bGuessType;

	if (m_bAllowTypeGuessing)
	{
		m_ulPropTag = PreferredStreamType(m_ulPropTag);

		// If we're guessing and happen to be handed PR_RTF_COMPRESSED, turn on our RTF editor
		if (PR_RTF_COMPRESSED == m_ulPropTag)
		{
			bEditPropAsRTF = true;
		}
	}

	m_bIsAB = bIsAB;
	m_bUseWrapEx = bUseWrapEx;
	m_ulRTFFlags = ulRTFFlags;
	m_ulInCodePage = ulInCodePage;
	m_ulOutCodePage = ulOutCodePage;
	m_ulStreamFlags = NULL;
	m_bDocFile = false;
	m_lpStream = NULL;
	m_StreamError = S_OK;
	m_bDisableSave = false;

	m_iTextBox = 0;
	if (bUseWrapEx)
	{
		m_iFlagBox = 1;
		m_iCodePageBox = 2;
		m_iBinBox = 3;
	}
	else
	{
		m_iFlagBox = 0xFFFFFFFF;
		m_iCodePageBox = 0xFFFFFFFF;
		m_iBinBox = 1;
	}

	UINT iNumBoxes = m_iBinBox + 1;

	m_iSmartViewBox = 0xFFFFFFFF;
	m_bDoSmartView = false;

	// Go ahead and open our property stream in case we decide to change our stream type
	// This way we can pick the right display elements
	OpenPropertyStream(false, bEditPropAsRTF);

	if (!m_bUseWrapEx && PT_BINARY == PROP_TYPE(m_ulPropTag))
	{
		m_bDoSmartView = true;
		m_iSmartViewBox = m_iBinBox + 1;
		iNumBoxes++;
	}

	if (bEditPropAsRTF) m_ulEditorType = EDITOR_RTF;
	else
	{
		switch (PROP_TYPE(m_ulPropTag))
		{
		case PT_STRING8:
			m_ulEditorType = EDITOR_STREAM_ANSI;
			break;
		case PT_UNICODE:
			m_ulEditorType = EDITOR_STREAM_UNICODE;
			break;
		case PT_BINARY:
		default:
			m_ulEditorType = EDITOR_STREAM_BINARY;
			break;
		}
	}

	CString szPromptPostFix;
	szPromptPostFix.Format(_T("\r\n%s"), (LPCTSTR)TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, false)); // STRING_OK

	SetPromptPostFix(szPromptPostFix);

	// Let's crack our property open and see what kind of controls we'll need for it
	// One control for text stream, one for binary
	CreateControls(iNumBoxes);
	InitPane(m_iTextBox, CreateCollapsibleTextPane(IDS_STREAMTEXT, false));
	if (bUseWrapEx)
	{
		InitPane(m_iFlagBox, CreateSingleLinePane(IDS_STREAMFLAGS, NULL, true));
		InitPane(m_iCodePageBox, CreateSingleLinePane(IDS_CODEPAGE, NULL, true));
	}

	InitPane(m_iBinBox, CreateCountedTextPane(IDS_STREAMBIN, false, IDS_CB));
	if (m_bDoSmartView)
	{
		InitPane(m_iSmartViewBox, CreateSmartViewPane(IDS_SMARTVIEW));
	}
} // CStreamEditor::CStreamEditor

CStreamEditor::~CStreamEditor()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpStream) m_lpStream->Release();
} // CStreamEditor::~CStreamEditor

// Used to call functions which need to be called AFTER controls are created
BOOL CStreamEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	ReadTextStreamFromProperty();

	if (m_bDoSmartView)
	{
		// Load initial smart view here
		SmartViewPane* lpSmartView = (SmartViewPane*)GetControl(m_iSmartViewBox);
		if (lpSmartView)
		{
			LPWSTR szSmartView = NULL;

			SPropValue sProp = { 0 };
			sProp.ulPropTag = CHANGE_PROP_TYPE(m_ulPropTag, PT_BINARY);
			if (GetBinaryUseControl(m_iBinBox, (size_t*)&sProp.Value.bin.cb, &sProp.Value.bin.lpb))
			{
				ULONG iStructType = InterpretPropSmartView(
					&sProp,
					m_lpMAPIProp,
					NULL,
					NULL,
					m_bIsAB,
					false,
					&szSmartView);

				lpSmartView->SetParser(iStructType);
				lpSmartView->SetStringW(szSmartView);

				delete[] szSmartView;
				szSmartView = NULL;
			}
			delete[] sProp.Value.bin.lpb;
		}
	}

	return bRet;
} // CStreamEditor::OnInitDialog

void CStreamEditor::OnOK()
{
	WriteTextStreamToProperty();
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
} // CStreamEditor::OnOK

void CStreamEditor::OpenPropertyStream(bool bWrite, bool bRTF)
{
	if (!m_lpMAPIProp) return;

	// Clear the previous stream if we have one
	if (m_lpStream) m_lpStream->Release();
	m_lpStream = NULL;

	HRESULT hRes = S_OK;
	LPSTREAM lpTmpStream = NULL;
	ULONG ulStgFlags = NULL;
	ULONG ulFlags = NULL;
	ULONG ulRTFFlags = m_ulRTFFlags;

	DebugPrintEx(DBGStream, CLASS, _T("OpenPropertyStream"), _T("opening property 0x%X (== %s) from %p, bWrite = 0x%X\n"), m_ulPropTag, (LPCTSTR)TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, true), m_lpMAPIProp, bWrite);

	if (bWrite)
	{
		ulStgFlags = STGM_READWRITE;
		ulFlags = MAPI_CREATE | MAPI_MODIFY;
		ulRTFFlags |= MAPI_MODIFY;

		if (m_bDocFile)
		{
			EC_MAPI(m_lpMAPIProp->OpenProperty(
				m_ulPropTag,
				&IID_IStreamDocfile,
				ulStgFlags,
				ulFlags,
				(LPUNKNOWN *)&lpTmpStream));
		}
		else
		{
			EC_MAPI(m_lpMAPIProp->OpenProperty(
				m_ulPropTag,
				&IID_IStream,
				ulStgFlags,
				ulFlags,
				(LPUNKNOWN *)&lpTmpStream));
		}
	}
	else
	{
		ulStgFlags = STGM_READ;
		ulFlags = NULL;
		WC_MAPI(m_lpMAPIProp->OpenProperty(
			m_ulPropTag,
			&IID_IStream,
			ulStgFlags,
			ulFlags,
			(LPUNKNOWN *)&lpTmpStream));

		// If we're guessing types, try again as a different type
		if (MAPI_E_NOT_FOUND == hRes && m_bAllowTypeGuessing)
		{
			ULONG ulPropTag = m_ulPropTag;
			switch (PROP_TYPE(ulPropTag))
			{
			case PT_STRING8:
			case PT_UNICODE:
				ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_BINARY);
				break;
			case PT_BINARY:
				ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_TSTRING);
				break;
			}

			if (ulPropTag != m_ulPropTag)
			{
				hRes = S_OK;
				DebugPrintEx(DBGStream, CLASS, _T("OpenPropertyStream"), _T("Retrying as 0x%X (== %s)\n"), m_ulPropTag, (LPCTSTR)TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, true));
				WC_MAPI(m_lpMAPIProp->OpenProperty(
					ulPropTag,
					&IID_IStream,
					ulStgFlags,
					ulFlags,
					(LPUNKNOWN *)&lpTmpStream));
				if (SUCCEEDED(hRes))
				{
					m_ulPropTag = ulPropTag;
				}
			}
		}

		// It's possible our stream was actually an docfile - give it a try
		if (FAILED(hRes))
		{
			hRes = S_OK;
			WC_MAPI(m_lpMAPIProp->OpenProperty(
				CHANGE_PROP_TYPE(m_ulPropTag, PT_OBJECT),
				&IID_IStreamDocfile,
				ulStgFlags,
				ulFlags,
				(LPUNKNOWN *)&lpTmpStream));
			if (SUCCEEDED(hRes))
			{
				m_bDocFile = true;
				m_ulPropTag = CHANGE_PROP_TYPE(m_ulPropTag, PT_OBJECT);
			}
		}
	}

	if (!bRTF)
	{
		m_lpStream = lpTmpStream;
		lpTmpStream = NULL;
	}
	else if (lpTmpStream)
	{
		EC_H(WrapStreamForRTF(
			lpTmpStream,
			m_bUseWrapEx,
			ulRTFFlags,
			m_ulInCodePage,
			m_ulOutCodePage,
			&m_lpStream,
			&m_ulStreamFlags));
	}

	if (lpTmpStream) lpTmpStream->Release();

	// Save off any error we got to display later
	m_StreamError = hRes;
}

void CStreamEditor::ReadTextStreamFromProperty()
{
	if (!m_lpMAPIProp) return;

	if (!IsValidEdit(m_iTextBox)) return;
	if (!IsValidEdit(m_iBinBox)) return;

	DebugPrintEx(DBGStream, CLASS, _T("ReadTextStreamFromProperty"), _T("opening property 0x%X (== %s) from %p\n"), m_ulPropTag, (LPCTSTR)TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, true), m_lpMAPIProp);

	// If we don't have a stream to display, put up an error instead
	if (FAILED(m_StreamError) || !m_lpStream)
	{
		CString szStreamErr;
		szStreamErr.FormatMessage(
			IDS_CANNOTOPENSTREAM,
			ErrorNameFromErrorCode(m_StreamError),
			m_StreamError);
		SetString(m_iTextBox, szStreamErr);
		SetEditReadOnly(m_iTextBox);
		SetEditReadOnly(m_iBinBox);
		return;
	}

	if (m_bUseWrapEx) // if m_bUseWrapEx, we're read only
	{
		SetEditReadOnly(m_iTextBox);
		SetEditReadOnly(m_iBinBox);
	}

	if (m_lpStream)
	{
		TextPane* lpPane = (TextPane*)GetControl(m_iBinBox);
		if (lpPane)
		{
			return lpPane->InitEditFromBinaryStream(m_lpStream);
		}
	}
} // CStreamEditor::ReadTextStreamFromProperty

// this will not work if we're using WrapCompressedRTFStreamEx
void CStreamEditor::WriteTextStreamToProperty()
{
	if (!IsValidEdit(m_iBinBox)) return;
	// If we couldn't get a read stream, we won't be able to get a write stream
	if (!m_lpStream) return;
	if (!m_lpMAPIProp) return;
	if (!IsDirty(m_iBinBox) && !IsDirty(m_iTextBox)) return; // If we didn't change it, don't write
	if (m_bUseWrapEx) return;

	HRESULT hRes = S_OK;

	// Reopen the property stream as writeable
	OpenPropertyStream(true, EDITOR_RTF == m_ulEditorType);

	// We started with a binary stream, pull binary back into the stream
	if (m_lpStream)
	{
		// We used to use EDITSTREAM here, but instead, just use GetBinaryUseControl and write it out.
		LPBYTE	lpb = NULL;
		size_t	cb = 0;
		ULONG cbWritten = 0;

		if (GetBinaryUseControl(m_iBinBox, &cb, &lpb))
		{
			EC_MAPI(m_lpStream->Write(lpb, (ULONG)cb, &cbWritten));
			DebugPrintEx(DBGStream, CLASS, _T("WriteTextStreamToProperty"), _T("wrote 0x%X\n"), cbWritten);

			EC_MAPI(m_lpStream->Commit(STGC_DEFAULT));

			if (m_bDisableSave)
			{
				DebugPrintEx(DBGStream, CLASS, _T("WriteTextStreamToProperty"), _T("Save was disabled.\n"));
			}
			else
			{
				EC_MAPI(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
			}
		}

		delete[] lpb;
	}

	DebugPrintEx(DBGStream, CLASS, _T("WriteTextStreamToProperty"), _T("Wrote out this stream:\n"));
	DebugPrintStream(DBGStream, m_lpStream);
} // CStreamEditor::WriteTextStreamToProperty

_Check_return_ ULONG CStreamEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if ((ULONG)-1 == i) return (ULONG)-1;

	LPBYTE	lpb = NULL;
	size_t	cb = 0;

	CountedTextPane* lpBinPane = (CountedTextPane*)GetControl(m_iBinBox);
	if (m_iTextBox == i && lpBinPane)
	{
		size_t cchStr = 0;
		LPSTR lpszA = NULL;
		LPWSTR lpszW = NULL;

		switch (m_ulEditorType)
		{
		case EDITOR_STREAM_ANSI:
		case EDITOR_RTF:
		case EDITOR_STREAM_BINARY:
		default:
			lpszA = GetEditBoxTextA(m_iTextBox, &cchStr);

			// What we just read includes a NULL terminator, in both the string and count.
			// When we write binary, we don't want to include this NULL
			if (cchStr) cchStr -= 1;

			lpBinPane->SetBinary((LPBYTE)lpszA, cchStr * sizeof(CHAR));
			lpBinPane->SetCount(cchStr * sizeof(CHAR));
			break;

		case EDITOR_STREAM_UNICODE:
			lpszW = GetEditBoxTextW(m_iTextBox, &cchStr);

			// What we just read includes a NULL terminator, in both the string and count.
			// When we write binary, we don't want to include this NULL
			if (cchStr) cchStr -= 1;

			lpBinPane->SetBinary((LPBYTE)lpszW, cchStr * sizeof(WCHAR));
			lpBinPane->SetCount(cchStr * sizeof(WCHAR));
			break;
		}
	}
	else if (m_iBinBox == i)
	{
		if (GetBinaryUseControl(m_iBinBox, &cb, &lpb))
		{
			switch (m_ulEditorType)
			{
			case EDITOR_STREAM_ANSI:
			case EDITOR_RTF:
			case EDITOR_STREAM_BINARY:
			default:
				// Treat as a NULL terminated string
				// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
				SetStringA(m_iTextBox, (LPCSTR)lpb, cb / sizeof(CHAR)+1);
				if (lpBinPane) lpBinPane->SetCount(cb);
				break;
			case EDITOR_STREAM_UNICODE:
				// Treat as a NULL terminated string
				// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
				SetStringW(m_iTextBox, (LPCWSTR)lpb, cb / sizeof(WCHAR)+1);
				if (lpBinPane) lpBinPane->SetCount(cb);
				break;
			}
		}
	}

	if (m_bDoSmartView)
	{
		SmartViewPane* lpSmartView = (SmartViewPane*)GetControl(m_iSmartViewBox);
		if (lpSmartView)
		{
			if (!cb && !lpb) (void)GetBinaryUseControl(m_iBinBox, &cb, &lpb);

			SBinary Bin = { 0 };
			Bin.cb = (ULONG)cb;
			Bin.lpb = lpb;

			lpSmartView->Parse(Bin);
		}
	}

	delete[] lpb;

	if (m_bUseWrapEx)
	{
		LPTSTR szFlags = NULL;
		InterpretFlags(flagStreamFlag, m_ulStreamFlags, &szFlags);
		SetStringf(m_iFlagBox, _T("0x%08X = %s"), m_ulStreamFlags, szFlags); // STRING_OK
		delete[] szFlags;
		szFlags = NULL;
		CString szTmp;
		szTmp.FormatMessage(IDS_CODEPAGES, m_ulInCodePage, m_ulOutCodePage);
		SetString(m_iCodePageBox, szTmp);
	}

	OnRecalcLayout();
	return i;
} // CStreamEditor::HandleChange

void CStreamEditor::SetEditReadOnly(ULONG iControl)
{
	if (IsValidEdit(iControl))
	{
		TextPane* lpPane = (TextPane*)GetControl(iControl);
		if (lpPane)
		{
			lpPane->SetEditReadOnly();
		}
	}
}

void CStreamEditor::DisableSave()
{
	m_bDisableSave = true;
}