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
		bool bIncludeABProps,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ CWnd* pParentWnd);
	virtual ~CPropertyTagEditor();

	_Check_return_ ULONG GetPropertyTag();

private:
	_Check_return_ ULONG HandleChange(UINT nID);
	void  OnEditAction1();
	void  OnEditAction2();
	_Check_return_ BOOL  OnInitDialog();
	void  PopulateFields(ULONG ulSkipField);
	_Check_return_ ULONG GetSelectedPropType();
	void  LookupNamedProp(ULONG ulSkipField, bool bCreate);

	ULONG		m_ulPropTag;
	bool		m_bIncludeABProps;
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
	_Check_return_ BOOL OnInitDialog();
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);
	void OnOK();

	ULONG		m_ulPropTag;
	bool		m_bIncludeABProps;
	LPMAPIPROP	m_lpMAPIProp;
};