#pragma once
// StreamEditor.h : header file

#include "Editor.h"

class CStreamEditor : public CEditor
{
public:
	CStreamEditor(
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
		ULONG ulOutCodePage);
	virtual ~CStreamEditor();

	void DisableSave();

private:
	BOOL  OnInitDialog();
	void  OpenPropertyStream(bool bWrite, bool bRTF);
	void  ReadTextStreamFromProperty();
	void  WriteTextStreamToProperty();
	_Check_return_ ULONG HandleChange(UINT nID);
	void  OnOK();
	void SetEditReadOnly(ULONG iControl);

	// source variables
	LPMAPIPROP	m_lpMAPIProp;
	LPSTREAM	m_lpStream;
	ULONG		m_ulPropTag;
	bool		m_bIsAB; // whether the tag is from the AB or not
	bool		m_bUseWrapEx;
	ULONG		m_ulRTFFlags;
	ULONG		m_ulInCodePage;
	ULONG		m_ulOutCodePage;
	ULONG		m_ulStreamFlags; // returned from WrapCompressedRTFStreamEx

	UINT		m_iTextBox;
	UINT		m_iFlagBox;
	UINT		m_iCodePageBox;
	UINT		m_iBinBox;
	UINT		m_iSmartViewBox;
	bool		m_bDoSmartView;
	bool		m_bDocFile;
	bool		m_bAllowTypeGuessing;
	bool		m_bDisableSave;
	ULONG		m_ulEditorType;
	HRESULT		m_StreamError;
};