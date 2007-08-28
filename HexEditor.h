#pragma once
// HexEditor.h : header file
//

#include "Editor.h"

class CHexEditor : public CEditor
{
public:
	CHexEditor(
		CWnd* pParentWnd);
	~CHexEditor();

	void InitAnsiString(
		LPCSTR szInputString);

	void InitUnicdeString(
		LPCWSTR szInputString);

	void InitBase64(
		LPCTSTR szInputBase64String);

	void InitBinary(
		LPSBinary lpsInputBin);

protected:
	//{{AFX_MSG(CPropertyEditor)
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
private:

	ULONG	HandleChange(UINT nID);
};
