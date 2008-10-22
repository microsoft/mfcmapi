#pragma once
// TagArrayEditor.h : header file

#include "Editor.h"

class CTagArrayEditor : public CEditor
{
public:
	CTagArrayEditor(
		CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		LPSPropTagArray lpTagArray,
		BOOL bIsAB,
		LPMAPIPROP	lpMAPIProp);
	virtual ~CTagArrayEditor();

	LPSPropTagArray DetachModifiedTagArray();

private:
	// Use this function to implement list editing
	BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);
	BOOL	OnInitDialog();
	void	OnOK();

	void	ReadTagArrayToList(ULONG ulListNum);
	void	WriteListToTagArray(ULONG ulListNum);

	// source variables
	LPSPropTagArray	m_lpTagArray;
	LPSPropTagArray	m_lpOutputTagArray;
	BOOL			m_bIsAB; // whether the tag is from the AB or not
	LPMAPIPROP		m_lpMAPIProp;
};