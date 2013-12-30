#pragma once
// PropertyEditor.h : header file

#include "Editor.h"

_Check_return_ HRESULT DisplayPropertyEditor(_In_ CWnd* pParentWnd,
											 UINT uidTitle,
											 UINT uidPrompt,
											 bool bIsAB,
											 _In_opt_ LPVOID lpAllocParent,
											 _In_opt_ LPMAPIPROP lpMAPIProp,
											 ULONG ulPropTag,
											 bool bMVRow,
											 _In_opt_ LPSPropValue lpsPropValue,
											 _Inout_opt_ LPSPropValue* lpNewValue);

class CPropertyEditor : public CEditor
{
public:
	CPropertyEditor(
		_In_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		bool bIsAB,
		bool bMVRow,
		_In_opt_ LPVOID lpAllocParent,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		_In_opt_ LPSPropValue lpsPropValue);
	virtual ~CPropertyEditor();

	// Get values after we've done the DisplayDialog
	_Check_return_ LPSPropValue DetachModifiedSPropValue();

private:
	BOOL  OnInitDialog();
	void  CreatePropertyControls();
	void  InitPropertyControls();
	void  WriteStringsToSPropValue();
	void  WriteSPropValueToObject();
	_Check_return_ ULONG HandleChange(UINT nID);
	void  OnOK();

	// source variables
	LPMAPIPROP		m_lpMAPIProp;
	ULONG			m_ulPropTag;
	bool			m_bIsAB; // whether the tag is from the AB or not
	LPSPropValue	m_lpsInputValue;
	LPSPropValue	m_lpsOutputValue;
	bool			m_bDirty;
	bool			m_bMVRow; // whether this row came from a multivalued property. Used for smart view parsing.
	SmartViewPane*  m_lpSmartView;

	// all calls to MAPIAllocateMore will use m_lpAllocParent
	// this is not something to be freed
	LPVOID			m_lpAllocParent;
}; // CPropertyEditor

class CMultiValuePropertyEditor : public CEditor
{
public:
	CMultiValuePropertyEditor(
		_In_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		bool bIsAB,
		_In_opt_ LPVOID lpAllocParent,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		_In_opt_ LPSPropValue lpsPropValue);
	virtual ~CMultiValuePropertyEditor();

	// Get values after we've done the DisplayDialog
	_Check_return_ LPSPropValue DetachModifiedSPropValue();

private:
	// Use this function to implement list editing
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);
	BOOL OnInitDialog();
	void CreatePropertyControls();
	void InitPropertyControls();
	void ReadMultiValueStringsFromProperty();
	void WriteSPropValueToObject();
	void WriteMultiValueStringsToSPropValue();
	void WriteMultiValueStringsToSPropValue(_In_ LPVOID lpParent, _In_ LPSPropValue lpsProp);
	void UpdateListRow(_In_ LPSPropValue lpProp, ULONG iMVCount);
	void UpdateSmartView();
	_Check_return_ ULONG HandleChange(UINT nID);
	void OnOK();

	// source variables
	LPMAPIPROP		m_lpMAPIProp;
	ULONG			m_ulPropTag;
	bool			m_bIsAB; // whether the tag is from the AB or not
	LPSPropValue	m_lpsInputValue;
	LPSPropValue	m_lpsOutputValue;

	// all calls to MAPIAllocateMore will use m_lpAllocParent
	// this is not something to be freed
	LPVOID			m_lpAllocParent;
}; // CMultiValuePropertyEditor