#pragma once
// TagArrayEditor.h : header file
//

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
	~CTagArrayEditor();

	LPSPropTagArray DetachModifiedTagArray();

protected:
	//{{AFX_MSG(CTagArrayEditor)
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	//Use this function to implement list editing
	virtual BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);

private:
	void ReadTagArrayToList(ULONG ulListNum);
	void WriteListToTagArray(ULONG ulListNum);
	//source variables
	LPSPropTagArray			m_lpTagArray;
	LPSPropTagArray			m_lpOutputTagArray;
	BOOL					m_bIsAB;//whether the tag is from the AB or not
	LPMAPIPROP				m_lpMAPIProp;

	BOOL	OnInitDialog();

	void	OnOK();
};
