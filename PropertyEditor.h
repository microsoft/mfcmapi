#pragma once
// PropertyEditor.h : header file

#include "Editor.h"

HRESULT DisplayPropertyEditor(CWnd* pParentWnd,
							  UINT uidTitle,
							  UINT uidPrompt,
							  BOOL bIsAB,
							  LPVOID lpAllocParent,
							  LPMAPIPROP lpMAPIProp,
							  ULONG ulPropTag,
							  LPSPropValue lpsPropValue,
							  LPSPropValue* lpNewValue);

class CPropertyEditor : public CEditor
{
public:
	CPropertyEditor(
		CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		BOOL bIsAB,
		LPVOID lpAllocParent,
		LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		LPSPropValue lpsPropValue);
	virtual ~CPropertyEditor();

	// Get values after we've done the DisplayDialog
	LPSPropValue DetachModifiedSPropValue();

private:
	BOOL	OnInitDialog();
	void	CreatePropertyControls();
	void	InitPropertyControls();
	void	WriteStringsToSPropValue();
	void	WriteSPropValueToObject();
	ULONG	HandleChange(UINT nID);
	void	OnOK();

	// source variables
	LPMAPIPROP		m_lpMAPIProp;
	ULONG			m_ulPropTag;
	BOOL			m_bIsAB; // whether the tag is from the AB or not
	LPSPropValue	m_lpsInputValue;
	LPSPropValue	m_lpsOutputValue;
	BOOL			m_bDirty;

	// all calls to MAPIAllocateMore will use m_lpAllocParent
	// this is not something to be freed
	LPVOID			m_lpAllocParent;
}; // CPropertyEditor

class CMultiValuePropertyEditor : public CEditor
{
public:
	CMultiValuePropertyEditor(
		CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		BOOL bIsAB,
		LPVOID lpAllocParent,
		LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		LPSPropValue lpsPropValue);
	virtual ~CMultiValuePropertyEditor();

	// Get values after we've done the DisplayDialog
	LPSPropValue DetachModifiedSPropValue();

private:
	// Use this function to implement list editing
	BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);
	BOOL	OnInitDialog();
	void	CreatePropertyControls();
	void	InitPropertyControls();
	void	ReadMultiValueStringsFromProperty(ULONG ulListNum);
	void	WriteSPropValueToObject();
	void	WriteMultiValueStringsToSPropValue(ULONG ulListNum);
	void	WriteMultiValueStringsToSPropValue(ULONG ulListNum, LPVOID lpParent, LPSPropValue lpsProp);
	void	UpdateListRow(LPSPropValue lpProp, ULONG ulListNum, ULONG iMVCount);
	void	UpdateSmartView(ULONG ulListNum);
	void	OnOK();

	// source variables
	LPMAPIPROP		m_lpMAPIProp;
	ULONG			m_ulPropTag;
	BOOL			m_bIsAB; // whether the tag is from the AB or not
	LPSPropValue	m_lpsInputValue;
	LPSPropValue	m_lpsOutputValue;
	BOOL			m_bDirty;

	// all calls to MAPIAllocateMore will use m_lpAllocParent
	// this is not something to be freed
	LPVOID			m_lpAllocParent;
}; // CMultiValuePropertyEditor