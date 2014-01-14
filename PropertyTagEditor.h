#pragma once
// PropertyTagEditor.h : header file

#include "Editor.h"

#define NOSKIPFIELD ((ULONG)0xffffffff)

class CPropertyTagEditor : public CEditor
{
public:
	CPropertyTagEditor(
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulPropTag,
		bool bIsAB,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_ CWnd* pParentWnd);
	virtual ~CPropertyTagEditor();

	_Check_return_ ULONG GetPropertyTag();

private:
	_Check_return_ ULONG HandleChange(UINT nID);
	void  OnEditAction1();
	void  OnEditAction2();
	BOOL  OnInitDialog();
	void  PopulateFields(ULONG ulSkipField);
	_Check_return_ ULONG GetSelectedPropType();
	void  LookupNamedProp(ULONG ulSkipField, bool bCreate);
	_Check_return_ CString GetDropStringUseControl(ULONG iControl);
	_Check_return_ int GetDropDownSelection(ULONG iControl);
	void InsertDropString(ULONG iControl, int iRow, _In_z_ LPCTSTR szText);
	void SetDropDownSelection(ULONG i, _In_opt_z_ LPCTSTR szText);

	ULONG		m_ulPropTag;
	bool		m_bIsAB;
	LPMAPIPROP	m_lpMAPIProp;
};

class CPropertySelector : public CEditor
{
public:
	CPropertySelector(
		bool bIncludeABProps,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ CWnd* pParentWnd);
	virtual ~CPropertySelector();

	_Check_return_ ULONG GetPropertyTag();

private:
	BOOL OnInitDialog();
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);
	void OnOK();
	_Check_return_ SortListData* GetSelectedListRowData(ULONG iControl);

	ULONG		m_ulPropTag;
	bool		m_bIncludeABProps;
	LPMAPIPROP	m_lpMAPIProp;
};