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
	virtual ~CRestrictEditor();

	LPSRestriction DetachModifiedSRestriction();

private:
	void	OnEditAction1();
	BOOL	OnInitDialog();
	ULONG	HandleChange(UINT nID);
	void	OnOK();

	LPSRestriction GetSourceRes();

	// source variables
	LPSRestriction			m_lpRes;
	LPVOID					m_lpAllocParent;

	// output variable
	LPSRestriction			m_lpOutputRes;
	BOOL					m_bModified;
};


class CCriteriaEditor : public CEditor
{
public:
	CCriteriaEditor(
		CWnd* pParentWnd,
		LPSRestriction lpRes,
		LPENTRYLIST lpEntryList,
		ULONG ulSearchState);
	virtual ~CCriteriaEditor();

	LPSRestriction DetachModifiedSRestriction();
	LPENTRYLIST DetachModifiedEntryList();
	ULONG GetSearchFlags();

private:
	// Use this function to implement list editing
	BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);
	void	OnEditAction1();
	BOOL	OnInitDialog();
	ULONG	HandleChange(UINT nID);
	void	InitListFromEntryList(ULONG ulListNum, LPENTRYLIST lpEntryList);
	void	OnOK();

	LPSRestriction GetSourceRes();

	LPSRestriction	m_lpSourceRes;
	LPSRestriction	m_lpNewRes;

	LPENTRYLIST		m_lpSourceEntryList;
	LPENTRYLIST		m_lpNewEntryList;

	ULONG			m_ulNewSearchFlags;
};