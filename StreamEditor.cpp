// StreamEditor.cpp : implementation file
//

#include "stdafx.h"
#include "StreamEditor.h"
#include "InterpretProp2.h"
#include "MAPIFunctions.h"
#include "ExtraPropTags.h"
#include "SmartView.h"

enum __StreamEditorTypes
{
	EDITOR_RTF,
	EDITOR_STREAM_BINARY,
	EDITOR_STREAM_ANSI,
	EDITOR_STREAM_UNICODE,
};

static TCHAR* CLASS = _T("CStreamEditor");

// Create an editor for a MAPI property - can be used to initialize stream and rtf stream editing as well
// Takes LPMAPIPROP and ulPropTag as input - will pull SPropValue from the LPMAPIPROP
CStreamEditor::CStreamEditor(
							 _In_ CWnd* pParentWnd,
							 UINT uidTitle,
							 UINT uidPrompt,
							 _In_ LPMAPIPROP lpMAPIProp,
							 ULONG ulPropTag,
							 BOOL bIsAB,
							 BOOL bEditPropAsRTF,
							 BOOL bUseWrapEx,
							 ULONG ulRTFFlags,
							 ULONG ulInCodePage,
							 ULONG ulOutCodePage):
CEditor(pParentWnd,uidTitle,uidPrompt,0,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_lpMAPIProp = lpMAPIProp;
	m_ulPropTag = ulPropTag;
	m_bIsAB = bIsAB;
	m_bUseWrapEx = bUseWrapEx;
	m_ulRTFFlags = ulRTFFlags;
	m_ulInCodePage = ulInCodePage;
	m_ulOutCodePage = ulOutCodePage;
	m_ulStreamFlags = NULL;
	m_bWriteAllowed = false;

	m_iTextBox = 0;
	m_iCBBox = 1;
	if (bUseWrapEx)
	{
		m_iFlagBox = 2;
		m_iCodePageBox = 3;
		m_iBinBox = 4;
	}
	else
	{
		m_iFlagBox = 0xFFFFFFFF;
		m_iCodePageBox = 0xFFFFFFFF;
		m_iBinBox = 2;
	}

	UINT iNumBoxes = m_iBinBox+1;

	m_iSmartViewBox = 0xFFFFFFFF;
	m_bDoSmartView = false;
	if (!m_bUseWrapEx && PT_BINARY == PROP_TYPE(m_ulPropTag))
	{
		m_bDoSmartView = true;
		m_iSmartViewBox = m_iBinBox+1;
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
	szPromptPostFix.Format(_T("\r\n%s"),(LPCTSTR) TagToString(m_ulPropTag,m_lpMAPIProp,m_bIsAB,false)); // STRING_OK

	SetPromptPostFix(szPromptPostFix);

	// Let's crack our property open and see what kind of controls we'll need for it
	// One control for text stream, one for binary
	CreateControls(iNumBoxes);
	InitMultiLine(m_iTextBox,IDS_STREAMTEXT ,NULL,false);
	InitSingleLine(m_iCBBox,IDS_CB,NULL,true);
	if (bUseWrapEx)
	{
		InitSingleLine(m_iFlagBox,IDS_STREAMFLAGS,NULL,true);
		InitSingleLine(m_iCodePageBox,IDS_CODEPAGE,NULL,true);
	}
	InitMultiLine(m_iBinBox,IDS_STREAMBIN,NULL,false);
	if (m_bDoSmartView)
	{
		InitMultiLine(m_iSmartViewBox,IDS_COLSMART_VIEW,NULL,true);
	}
} // CStreamEditor::CStreamEditor

CStreamEditor::~CStreamEditor()
{
	TRACE_DESTRUCTOR(CLASS);
} // CStreamEditor::~CStreamEditor

// Used to call functions which need to be called AFTER controls are created
_Check_return_ BOOL CStreamEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	ReadTextStreamFromProperty();

	return bRet;
} // CStreamEditor::OnInitDialog

void CStreamEditor::OnOK()
{
	WriteTextStreamToProperty();
	CDialog::OnOK(); // don't need to call CEditor::OnOK
} // CStreamEditor::OnOK

void CStreamEditor::ReadTextStreamFromProperty()
{
	if (!m_lpMAPIProp) return;

	if (!IsValidEdit(m_iTextBox)) return;
	if (!IsValidEdit(m_iBinBox)) return;

	HRESULT hRes = S_OK;
	LPSTREAM lpTmpStream = NULL;
	LPSTREAM lpTmpRTFStream = NULL;
	LPSTREAM lpStreamIn = NULL;

	DebugPrintEx(DBGStream,CLASS,_T("ReadTextStreamFromProperty"),_T("opening property 0x%X (== %s) from %p\n"),m_ulPropTag,(LPCTSTR) TagToString(m_ulPropTag,m_lpMAPIProp,m_bIsAB,true),m_lpMAPIProp);

	WC_H(m_lpMAPIProp->OpenProperty(
		m_ulPropTag,
		&IID_IStream,
		STGM_READ,
		0,
		(LPUNKNOWN *)&lpTmpStream));
	if (MAPI_E_NOT_FOUND == hRes)
	{
		WARNHRESMSG(hRes,IDS_PROPERTYNOTFOUND);
		return;
	}
	if (FAILED(hRes) || !lpTmpStream)
	{
		CString szStreamErr;
		szStreamErr.FormatMessage(
			IDS_CANNOTOPENSTREAM,
			ErrorNameFromErrorCode(hRes),
			hRes);
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
	if (EDITOR_RTF != m_ulEditorType)
	{
		lpStreamIn = lpTmpStream;
	}
	else
	{
		EC_H(WrapStreamForRTF(
			lpTmpStream,
			m_bUseWrapEx,
			m_ulRTFFlags,
			m_ulInCodePage,
			m_ulOutCodePage,
			&lpTmpRTFStream,
			&m_ulStreamFlags));

		lpStreamIn = lpTmpRTFStream;
	}
	if (lpStreamIn)
	{
		InitEditFromBinaryStream(m_iBinBox,lpStreamIn);
	}

	if (!m_bUseWrapEx) m_bWriteAllowed = true;

	if (lpTmpRTFStream) lpTmpRTFStream->Release();
	if (lpTmpStream) lpTmpStream->Release();
} // CStreamEditor::ReadTextStreamFromProperty

// this will not work if we're using WrapCompressedRTFStreamEx
void CStreamEditor::WriteTextStreamToProperty()
{
	if (!IsValidEdit(m_iBinBox)) return;
	if (!m_bWriteAllowed) return; // don't want to write when we're in an error state
	if (!m_lpMAPIProp) return;
	if (!EditDirty(m_iBinBox) && !EditDirty(m_iTextBox)) return; // If we didn't change it, don't write
	if (m_bUseWrapEx) return;

	HRESULT hRes = S_OK;
	LPSTREAM lpTmpStream = NULL;
	LPSTREAM lpTmpRTFStream = NULL;
	LPSTREAM lpStreamOut = NULL;

	EC_H(m_lpMAPIProp->OpenProperty(
		m_ulPropTag,
		&IID_IStream,
		STGM_READWRITE,
		MAPI_CREATE	| MAPI_MODIFY,
		(LPUNKNOWN *)&lpTmpStream));

	if (FAILED(hRes) || !lpTmpStream) return;

	if (EDITOR_RTF != m_ulEditorType)
	{
		lpStreamOut = lpTmpStream;
	}
	else
	{
		EC_H(WrapStreamForRTF(
			lpTmpStream,
			m_bUseWrapEx,
			m_ulRTFFlags | MAPI_MODIFY,
			m_ulInCodePage,
			m_ulOutCodePage,
			&lpTmpRTFStream,
			&m_ulStreamFlags));

		lpStreamOut = lpTmpRTFStream;
	}

	// We started with a binary stream, pull binary back into the stream
	if (lpStreamOut)
	{
		// We used to use EDITSTREAM here, but instead, just use GetBinaryUseControl and write it out.
		LPBYTE	lpb = NULL;
		size_t	cb = 0;
		ULONG cbWritten = 0;

		if (GetBinaryUseControl(m_iBinBox, &cb, &lpb))
		{
			EC_H(lpStreamOut->Write(lpb, (ULONG) cb, &cbWritten));
			DebugPrintEx(DBGStream,CLASS,_T("WriteTextStreamToProperty"),_T("wrote 0x%X\n"),cbWritten);

			EC_H(lpStreamOut->Commit(STGC_DEFAULT));

			EC_H(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
		}
		delete[] lpb;
	}

	DebugPrintEx(DBGStream,CLASS, _T("WriteTextStreamToProperty"),_T("Wrote out this stream:\n"));
	DebugPrintStream(DBGStream,lpStreamOut);

	if (lpTmpRTFStream) lpTmpRTFStream->Release();
	if (lpTmpStream) lpTmpStream->Release();
} // CStreamEditor::WriteTextStreamToProperty

_Check_return_ ULONG CStreamEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if ((ULONG) -1 == i) return (ULONG) -1;

	LPBYTE	lpb = NULL;
	size_t	cb = 0;

	if (m_iTextBox == i)
	{
		size_t cchStr = 0;
		switch (m_ulEditorType)
		{
		case EDITOR_STREAM_ANSI:
		case EDITOR_RTF:
		case EDITOR_STREAM_BINARY:
		default:
			{
				LPSTR lpszA = GetEditBoxTextA(m_iTextBox, &cchStr);

				// What we just read includes a NULL terminator, in both the string and count.
				// When we write binary, we don't want to include this NULL
				if (cchStr) cchStr -= 1;

				SetBinary(m_iBinBox, (LPBYTE) lpszA, cchStr * sizeof(CHAR));
				SetSize(m_iCBBox, cchStr * sizeof(CHAR));
				break;
			}
		case EDITOR_STREAM_UNICODE:
			{
				LPWSTR lpszW = GetEditBoxTextW(m_iTextBox, &cchStr);

				// What we just read includes a NULL terminator, in both the string and count.
				// When we write binary, we don't want to include this NULL
				if (cchStr) cchStr -= 1;

				SetBinary(m_iBinBox, (LPBYTE) lpszW, cchStr * sizeof(WCHAR));
				SetSize(m_iCBBox, cchStr * sizeof(WCHAR));
				delete[] lpszW;
				break;
			}
		}
	}
	else if (m_iBinBox == i)
	{
		if (GetBinaryUseControl(m_iBinBox,&cb,&lpb))
		{
			switch (m_ulEditorType)
			{
			case EDITOR_STREAM_ANSI:
			case EDITOR_RTF:
			case EDITOR_STREAM_BINARY:
			default:
				// Treat as a NULL terminated string
				// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
				SetStringA(m_iTextBox,(LPCSTR) lpb, cb/sizeof(CHAR)+1);
				SetSize(m_iCBBox, cb);
				break;
			case EDITOR_STREAM_UNICODE:
				// Treat as a NULL terminated string
				// GetBinaryUseControl includes extra NULLs at the end of the buffer to make this work
				SetStringW(m_iTextBox,(LPCWSTR) lpb, cb/sizeof(WCHAR)+1);
				SetSize(m_iCBBox, cb);
				break;
			}
		}
	}

	if (m_bDoSmartView)
	{
		if (!cb && !lpb) (void) GetBinaryUseControl(m_iBinBox,&cb,&lpb);

		LPTSTR szSmartView = NULL;
		SBinary Bin = {0};
		Bin.cb = (ULONG) cb;
		Bin.lpb = lpb;
		SPropValue sProp = {0};
		sProp.ulPropTag = m_ulPropTag;
		sProp.Value.bin = Bin;

		InterpretPropSmartView(
			&sProp,
			NULL,
			NULL,
			NULL,
			false,
			&szSmartView);
		SetString(m_iSmartViewBox,szSmartView);
		delete[] szSmartView;
	}
	delete[] lpb;

	if (m_bUseWrapEx)
	{
		LPTSTR szFlags = NULL;
		InterpretFlags(flagStreamFlag, m_ulStreamFlags, &szFlags);
		SetStringf(m_iFlagBox,_T("0x%08X = %s"),m_ulStreamFlags,szFlags); // STRING_OK
		delete[] szFlags;
		szFlags = NULL;
		CString szTmp;
		szTmp.FormatMessage(IDS_CODEPAGES,m_ulInCodePage, m_ulOutCodePage);
		SetString(m_iCodePageBox,szTmp);
	}

	return i;
} // CStreamEditor::HandleChange