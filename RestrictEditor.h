#pragma once
// RestrictEditor.h : header file
//

#include "Editor.h"

class CRestrictEditor : public CEditor
{
public:
	CRestrictEditor(
		CWnd* pParentWnd,
		LPVOID lpAllocParent,
		LPSRestriction lpRes);
	~CRestrictEditor();

	LPSRestriction DetachModifiedSRestriction();
	virtual void OnEditAction1();

protected:
	//{{AFX_MSG(CRestrictEditor)
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

private:
	//source variables
	LPSRestriction			m_lpRes;
	LPVOID					m_lpAllocParent;

	//output variable
	LPSRestriction			m_lpOutputRes;
	BOOL					m_bModified;

	BOOL	OnInitDialog();
	ULONG	HandleChange(UINT nID);

	LPSRestriction GetSourceRes();

	void	OnOK();
};


class CCriteriaEditor : public CEditor
{
public:
	CCriteriaEditor(
		CWnd* pParentWnd,
		LPSRestriction lpRes,
		LPENTRYLIST lpEntryList,
		ULONG ulSearchState);
	~CCriteriaEditor();
	virtual void OnEditAction1();
	LPSRestriction DetachModifiedSRestriction();
	LPENTRYLIST DetachModifiedEntryList();
	ULONG GetSearchFlags();
protected:
	//Use this function to implement list editing
	virtual BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);

private:
	BOOL	OnInitDialog();
	ULONG	HandleChange(UINT nID);
	void	InitListFromEntryList(ULONG ulListNum, LPENTRYLIST lpEntryList);
	LPSRestriction GetSourceRes();
	void	OnOK();

	LPSRestriction	m_lpSourceRes;
	LPSRestriction	m_lpNewRes;

	LPENTRYLIST		m_lpSourceEntryList;
	LPENTRYLIST		m_lpNewEntryList;

	ULONG			m_ulNewSearchFlags;
};
