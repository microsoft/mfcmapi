#pragma once
// PropertyTagEditor.h : header file

#include "Editor.h"

#define NOSKIPFIELD ((ULONG)0xffffffff)

class CPropertyTagEditor : public CEditor
{
public:
	CPropertyTagEditor(
		UINT		uidTitle,
		UINT		uidPrompt,
		ULONG		ulPropTag,
		BOOL		bIncludeABProps,
		LPMAPIPROP	lpMAPIProp,
		CWnd*		pParentWnd);
	virtual ~CPropertyTagEditor();

	ULONG GetPropertyTag();

private:
	ULONG	HandleChange(UINT nID);
	void	OnEditAction1();
	void	OnEditAction2();
	BOOL	OnInitDialog();
	void	PopulateFields(ULONG ulSkipField);
	BOOL	GetSelectedGUID(LPGUID lpSelectedGUID);
	BOOL	GetSelectedPropType(ULONG* ulPropType);
	void	LookupNamedProp(ULONG ulSkipField, BOOL bCreate);

	ULONG		m_ulPropTag;
	BOOL		m_bIncludeABProps;
	LPMAPIPROP	m_lpMAPIProp;
};

class CPropertySelector : public CEditor
{
public:
	CPropertySelector(
		BOOL		bIncludeABProps,
		LPMAPIPROP	lpMAPIProp,
		CWnd*		pParentWnd);
	virtual ~CPropertySelector();

	ULONG GetPropertyTag();

private:
	BOOL	OnInitDialog();
	BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);
	void	OnOK();

	ULONG		m_ulPropTag;
	BOOL		m_bIncludeABProps;
	LPMAPIPROP	m_lpMAPIProp;
};