#pragma once
// TagArrayEditor.h : header file

#include "Editor.h"

class CTagArrayEditor : public CEditor
{
public:
	CTagArrayEditor(
		_In_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		_In_opt_ LPSPropTagArray lpTagArray,
		BOOL bIsAB,
		_In_ LPMAPIPROP	lpMAPIProp);
	virtual ~CTagArrayEditor();

	_Check_return_ LPSPropTagArray DetachModifiedTagArray();

private:
	// Use this function to implement list editing
	_Check_return_ BOOL DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);
	_Check_return_ BOOL OnInitDialog();
	void OnOK();

	void ReadTagArrayToList(ULONG ulListNum);
	void WriteListToTagArray(ULONG ulListNum);

	// source variables
	LPSPropTagArray	m_lpTagArray;
	LPSPropTagArray	m_lpOutputTagArray;
	BOOL			m_bIsAB; // whether the tag is from the AB or not
	LPMAPIPROP		m_lpMAPIProp;
};