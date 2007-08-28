#pragma once
// StreamEditor.h : header file
//

#include "Editor.h"

class CStreamEditor : public CEditor
{
public:
	CStreamEditor(
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
		ULONG ulOutCodePage);
	~CStreamEditor();

protected:
	//{{AFX_MSG(CStreamEditor)
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

private:
	ULONG					m_ulEditorType;

	//source variables
	LPMAPIPROP				m_lpMAPIProp;
	ULONG					m_ulPropTag;
	BOOL					m_bIsAB;//whether the tag is from the AB or not
	BOOL					m_bUseWrapEx;
	ULONG					m_ulRTFFlags;
	ULONG					m_ulInCodePage;
	ULONG					m_ulOutCodePage;
	ULONG					m_ulStreamFlags;//returned from WrapCompressedRTFStreamEx

	UINT					m_iTextBox;
	UINT					m_iCBBox;
	UINT					m_iFlagBox;
	UINT					m_iCodePageBox;
	UINT					m_iBinBox;
	BOOL					m_bWriteAllowed;

	BOOL	OnInitDialog();
	void	ReadTextStreamFromProperty();
	void	WriteTextStreamToProperty();
	ULONG	HandleChange(UINT nID);

	void	OnOK();
};
