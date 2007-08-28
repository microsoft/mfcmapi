// StreamEditor.cpp : implementation file
//

#include "stdafx.h"
#include "Error.h"

#include "StreamEditor.h"

#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "MAPIFunctions.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

enum __StreamEditorTypes
{
	EDITOR_RTF,
	EDITOR_STREAM,
};

static TCHAR* CLASS = _T("CStreamEditor");

//Create an editor for a MAPI property - can be used to initialize stream and rtf stream editing as well
//Takes LPMAPIPROP and ulPropTag as input - will pull SPropValue from the LPMAPIPROP
CStreamEditor::CStreamEditor(
								 CWnd* pParentWnd,
								 UINT uidTitle,
								 UINT uidPrompt,
								 LPMAPIPROP lpMAPIProp,
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

	if (bEditPropAsRTF) m_ulEditorType = EDITOR_RTF;
	else m_ulEditorType = EDITOR_STREAM;

	CString szPromptPostFix;
	szPromptPostFix.Format(_T("\r\n%s"),TagToString(m_ulPropTag,m_lpMAPIProp,m_bIsAB,false));// STRING_OK

	SetPromptPostFix(szPromptPostFix);

	//Let's crack our property open and see what kind of controls we'll need for it
	//One control for text stream, one for binary
	CreateControls(m_iBinBox+1);
	InitMultiLine(m_iTextBox,IDS_STREAMTEXT ,NULL,false);
	InitSingleLine(m_iCBBox,IDS_CB,NULL,true);
	if (bUseWrapEx)
	{
		InitSingleLine(m_iFlagBox,IDS_STREAMFLAGS,NULL,true);
		InitSingleLine(m_iCodePageBox,IDS_CODEPAGE,NULL,true);
	}
	InitMultiLine(m_iBinBox,IDS_STREAMBIN,NULL,false);
}

CStreamEditor::~CStreamEditor()
{
	TRACE_DESTRUCTOR(CLASS);
}

BEGIN_MESSAGE_MAP(CStreamEditor, CEditor)
//{{AFX_MSG_MAP(CStreamEditor)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

//Used to call functions which need to be called AFTER controls are created
BOOL CStreamEditor::OnInitDialog()
{
	HRESULT hRes = S_OK;

	EC_B(CEditor::OnInitDialog());

	ReadTextStreamFromProperty();

	return HRES_TO_BOOL(hRes);
}

void CStreamEditor::OnOK()
{
	WriteTextStreamToProperty();
	CDialog::OnOK();//don't need to call CEditor::OnOK
}

class MyRichEditCookie
{
public:
	LPSTREAM  m_pData;
	MyRichEditCookie(LPSTREAM pData)
	{
		m_pData = pData;
	}
};

static DWORD CALLBACK EditStreamWriteCallBack(
											  DWORD_PTR dwCookie,
											  LPBYTE pbBuff,
											  LONG cb,
											  LONG *pcb)
{
	HRESULT hRes = S_OK;
	if (!pbBuff || !pcb || !dwCookie) return 0;

	MyRichEditCookie* pCookie = (MyRichEditCookie*) dwCookie;

	LPSTREAM stmData = pCookie->m_pData;
	ULONG cbWritten = 0;

	*pcb = 0;

	DebugPrint(DBGGeneric,_T("EditStreamWriteCallBack:cb = %d\n"),cb);

	EC_H(stmData->Write(pbBuff,cb,&cbWritten));

	DebugPrint(DBGGeneric,_T("EditStreamWriteCallBack: wrote %d bytes\n"),cbWritten);

	*pcb = cbWritten;

	return 0;
}

static DWORD CALLBACK EditStreamReadCallBack(
											  DWORD_PTR dwCookie,
											  LPBYTE pbBuff,
											  LONG cb,
											  LONG *pcb)
{
	HRESULT hRes = S_OK;
	if (!pbBuff || !pcb || !dwCookie) return 0;

	MyRichEditCookie* pCookie = (MyRichEditCookie*) dwCookie;

	LPSTREAM stmData = pCookie->m_pData;
	ULONG cbRead = 0;

	*pcb = 0;

	DebugPrint(DBGGeneric,_T("EditStreamWriteCallBack:cb = %d\n"),cb);

	EC_H(stmData->Read(pbBuff,cb,&cbRead));

	DebugPrint(DBGGeneric,_T("EditStreamReadCallBack: read %d bytes\n"),cbRead);

	*pcb = cbRead;

	return 0;
}

void CStreamEditor::ReadTextStreamFromProperty()
{
	if (!m_lpMAPIProp) return;
	if (!m_lpControls) return;

	if (!IsValidEdit(m_iTextBox)) return;
	if (!IsValidEdit(m_iBinBox)) return;

	HRESULT hRes = S_OK;
	LPSTREAM lpTmpStream = NULL;
	LPSTREAM lpTmpRTFStream = NULL;
	LPSTREAM lpStreamIn = NULL;

	DebugPrintEx(DBGGeneric,CLASS,_T("ReadTextStreamFromProperty"),_T("opening property 0x%X (==%s) from 0x%X\n"),m_ulPropTag,(LPCTSTR) TagToString(m_ulPropTag,m_lpMAPIProp,m_bIsAB,true),m_lpMAPIProp);

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
	if (MAPI_E_NO_SUPPORT == hRes || !lpTmpStream)
	{
		WARNHRESMSG(hRes,IDS_STREAMNOTSUPPORTED);
		LoadString(m_iTextBox, IDS_STREAMNOTSUPPORTED);
		m_lpControls[m_iTextBox].UI.lpEdit->EditBox.SetBackgroundColor(false,GetSysColor(COLOR_BTNFACE));
		m_lpControls[m_iTextBox].UI.lpEdit->EditBox.SetReadOnly();
		m_lpControls[m_iBinBox].UI.lpEdit->EditBox.SetBackgroundColor(false,GetSysColor(COLOR_BTNFACE));
		m_lpControls[m_iBinBox].UI.lpEdit->EditBox.SetReadOnly();
		return;
	}
	if (m_bUseWrapEx)//if m_bUseWrapEx, we're read only
	{
		m_lpControls[m_iTextBox].UI.lpEdit->EditBox.SetBackgroundColor(false,GetSysColor(COLOR_BTNFACE));
		m_lpControls[m_iTextBox].UI.lpEdit->EditBox.SetReadOnly();
		m_lpControls[m_iBinBox].UI.lpEdit->EditBox.SetBackgroundColor(false,GetSysColor(COLOR_BTNFACE));
		m_lpControls[m_iBinBox].UI.lpEdit->EditBox.SetReadOnly();
	}
	if (EDITOR_STREAM == m_ulEditorType)
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
		EDITSTREAM			es = {0, 0, EditStreamReadCallBack};
		MyRichEditCookie	cookie(lpStreamIn);
		UINT				uFormat = SF_TEXT;

		//If we're unicode, we need to pass this flag to use Unicode RichEdit
		//However, PR_RTF_COMPRESSED don't play Unicode, so we check if we're RTF first
		if (PROP_TYPE(m_ulPropTag) == PT_UNICODE)
		{
			if (EDITOR_STREAM == m_ulEditorType)
				uFormat |= SF_UNICODE;
		}

		long				lBytesRead = 0;

		es.dwCookie = (DWORD_PTR)&cookie;

		//read the 'text' stream into one control
		lBytesRead = m_lpControls[m_iTextBox].UI.lpEdit->EditBox.StreamIn(uFormat,es);
		DebugPrintEx(DBGGeneric,CLASS,_T("ReadTextStreamFromProperty"),_T("read %d bytes from the stream\n"),lBytesRead);
	}

	if (!m_bUseWrapEx) m_bWriteAllowed = true;

	//cheat to get the rest filled out
	HandleChange(m_lpControls[m_iTextBox].nID);
	if (lpTmpRTFStream) lpTmpRTFStream->Release();
	if (lpTmpStream) lpTmpStream->Release();
}

//this will not work if we're using WrapCompressedRTFStreamEx
void CStreamEditor::WriteTextStreamToProperty()
{
	if (!IsValidEdit(m_iTextBox)) return;
	if (!m_bWriteAllowed) return;//don't want to write when we're in an error state
	if (!m_lpMAPIProp) return;
	if (!m_lpControls[m_iTextBox].UI.lpEdit->EditBox.GetModify()) return;
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

	if (EDITOR_STREAM == m_ulEditorType)
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

	//We started with a text stream, pull text back into the stream
	if (lpStreamOut)
	{
		EDITSTREAM			es = {0, 0, EditStreamWriteCallBack};
		MyRichEditCookie	cookie(lpStreamOut);
		LONG				cb = 0;
		UINT				uFormat = SF_TEXT;

		//If we're unicode, we need to pass this flag to use Unicode RichEdit
		//However, PR_RTF_COMPRESSED don't play Unicode, so we check if we're RTF first
		if (PROP_TYPE(m_ulPropTag) == PT_UNICODE)
		{
			if (EDITOR_STREAM == m_ulEditorType)
				uFormat |= SF_UNICODE;
		}

		es.dwCookie	= (DWORD_PTR)&cookie;

		cb = m_lpControls[m_iTextBox].UI.lpEdit->EditBox.StreamOut(uFormat,es);
		DebugPrintEx(DBGGeneric,CLASS,_T("WriteTextStreamToProperty"),_T("wrote 0x%X\n"),cb);

		EC_H(lpStreamOut->Commit(STGC_DEFAULT));

		EC_H(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
	}

	DebugPrintEx(DBGGeneric,CLASS, _T("WriteTextStreamToProperty"),_T("Wrote out this stream:\n"));
	DebugPrintStream(DBGGeneric,lpStreamOut);

	if (lpTmpRTFStream) lpTmpRTFStream->Release();
	if (lpTmpStream) lpTmpStream->Release();
}

ULONG CStreamEditor::HandleChange(UINT nID)
{
	if (!m_lpControls) return (ULONG) -1;
	ULONG i = CEditor::HandleChange(nID);

	if ((ULONG) -1 == i) return (ULONG) -1;
	if (m_iTextBox == i)
	{
		LPSTR lpszA = GetEditBoxTextA(m_iTextBox);

		//cchszW was count of characters in the szW, including the null terminator
		//count of bytes in the lpszA is then cchszW-1
		size_t cbStr = m_lpControls[m_iTextBox].UI.lpEdit->cchszW-1;

		SetBinary(2, (LPBYTE) lpszA, cbStr);

		SetSize(m_iCBBox, cbStr);
	}
	else if (m_iBinBox == i)
	{
		LPBYTE	lpb = NULL;
		size_t	cb = 0;

		if (GetBinaryUseControl(m_iBinBox,&cb,&lpb))
		{
			SetStringA(0,(LPCSTR) lpb);
			SetSize(m_iCBBox, cb);
		}

		delete[] lpb;
	}

	if (m_bUseWrapEx)
	{
		HRESULT hRes = S_OK;
		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagStreamFlag, m_ulStreamFlags, &szFlags));
		SetStringf(m_iFlagBox,_T("0x%08X = %s"),m_ulStreamFlags,szFlags);// STRING_OK
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
		CString szTmp;
		szTmp.FormatMessage(IDS_CODEPAGES,m_ulInCodePage, m_ulOutCodePage);
		SetString(m_iCodePageBox,szTmp);
	}

	return i;
}