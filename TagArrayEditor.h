#pragma once
// TagArrayEditor.h : header file

#include "Editor.h"

class CTagArrayEditor : public CEditor
{
public:
	CTagArrayEditor(
		_In_opt_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		_In_opt_ LPMAPITABLE lpContentsTable,
		_In_opt_ LPSPropTagArray lpTagArray,
		bool bIsAB,
		_In_opt_ LPMAPIPROP	lpMAPIProp);
	virtual ~CTagArrayEditor();

	_Check_return_ LPSPropTagArray DetachModifiedTagArray();

private:
	// Use this function to implement list editing
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);
	BOOL OnInitDialog();
	void OnOK();
	void OnEditAction1();
	void OnEditAction2();

	void ReadTagArrayToList(ULONG ulListNum);
	void WriteListToTagArray(ULONG ulListNum);

	// source variables
	LPMAPITABLE m_lpContentsTable;
	LPSPropTagArray	m_lpTagArray;
	LPSPropTagArray	m_lpOutputTagArray;
	bool			m_bIsAB; // whether the tag is from the AB or not
	LPMAPIPROP		m_lpMAPIProp;
	ULONG m_ulSetColumnsFlags;
};