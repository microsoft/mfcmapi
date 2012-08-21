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
		bool bIsAB,
		bool bEditPropAsRTF,
		bool bUseWrapEx,
		ULONG ulRTFFlags,
		ULONG ulInCodePage,
		ULONG ulOutCodePage);
	virtual ~CStreamEditor();

private:
	_Check_return_ BOOL  OnInitDialog();
	void  ReadTextStreamFromProperty();
	void  WriteTextStreamToProperty();
	_Check_return_ ULONG HandleChange(UINT nID);
	void  OnOK();

	// source variables
	LPMAPIPROP	m_lpMAPIProp;
	ULONG		m_ulPropTag;
	bool		m_bIsAB; // whether the tag is from the AB or not
	bool		m_bUseWrapEx;
	ULONG		m_ulRTFFlags;
	ULONG		m_ulInCodePage;
	ULONG		m_ulOutCodePage;
	ULONG		m_ulStreamFlags; // returned from WrapCompressedRTFStreamEx

	UINT		m_iTextBox;
	UINT		m_iCBBox;
	UINT		m_iFlagBox;
	UINT		m_iCodePageBox;
	UINT		m_iBinBox;
	UINT		m_iSmartViewBox;
	bool		m_bWriteAllowed;
	bool		m_bDoSmartView;
	bool		m_bDocFile;
	ULONG		m_ulEditorType;
};