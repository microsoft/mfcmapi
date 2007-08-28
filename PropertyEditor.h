#pragma once
// PropertyEditor.h : header file
//

#include "Editor.h"

class CPropertyEditor : public CEditor
{
public:
	CPropertyEditor(
		CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		BOOL bIsAB,
		LPVOID lpAllocParent);
	~CPropertyEditor();

	void InitPropValue(
		LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag);

	void InitPropValue(
		LPSPropValue lpsPropValue);

	//Get values after we've done the DisplayDialog
	LPSPropValue DetachModifiedSPropValue();

protected:
	//{{AFX_MSG(CPropertyEditor)
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	//Use this function to implement list editing
	virtual BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);

	BOOL	OnInitDialog();
private:

	void Constructor();
	ULONG					m_ulEditorType;

	//source variables
	LPMAPIPROP				m_lpMAPIProp;
	ULONG					m_ulPropTag;
	BOOL					m_bIsAB;//whether the tag is from the AB or not
	LPSPropValue			m_lpsInputValue;
	LPSPropValue			m_lpsOutputValue;
	BOOL					m_bShouldFreeOutputValue;
	BOOL					m_bReadOnly;
	//all calls to MAPIAllocateMore will use m_lpAllocParent
	//this is not something to be freed
	LPVOID					m_lpAllocParent;

	void	CreatePropertyControls();
	void	InitPropertyControls();
	void	ReadMultiValueStringsFromProperty(ULONG ulListNum);
	void	WriteStringsToSPropValue();
	void	WriteSPropValueToObject();
	void	WriteMultiValueStringsToSPropValue(ULONG ulListNum);
	ULONG	HandleChange(UINT nID);

	void	OnOK();
};
