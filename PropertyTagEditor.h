#pragma once
// PropertyTagEditor.h : header file
//

#include "Editor.h"

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
	~CPropertyTagEditor();

	ULONG GetPropertyTag();

protected:
	//{{AFX_MSG(CPropertyEditor)
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	BOOL	OnInitDialog();
private:
	ULONG	HandleChange(UINT nID);
	void	PopulateFields(ULONG ulSkipField);
	virtual void OnEditAction1();
	virtual void OnEditAction2();
	BOOL	GetSelectedGUID(LPGUID lpSelectedGUID);
	BOOL	GetSelectedPropType(ULONG* ulPropType);

	ULONG		m_ulPropTag;
	BOOL		m_bIncludeABProps;
	LPMAPIPROP	m_lpMAPIProp;
};

#define NOSKIPFIELD ((ULONG)0xffffffff)

class CPropertySelector : public CEditor
{
public:
	CPropertySelector(
		BOOL		bIncludeABProps,
		LPMAPIPROP	lpMAPIProp,
		CWnd*		pParentWnd);
	~CPropertySelector();

	ULONG GetPropertyTag();

protected:
	//{{AFX_MSG(CPropertyEditor)
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	BOOL	OnInitDialog();
	virtual BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);
private:
	void	OnOK();

	ULONG		m_ulPropTag;
	BOOL		m_bIncludeABProps;
	LPMAPIPROP	m_lpMAPIProp;
};
