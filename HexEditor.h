#pragma once
// HexEditor.h : header file

#include "Editor.h"

class CHexEditor : public CEditor
{
public:
	CHexEditor(
		CWnd* pParentWnd);
	virtual ~CHexEditor();

	void InitAnsiString(
		LPCSTR szInputString);

	void InitUnicodeString(
		LPCWSTR szInputString);

	void InitBase64(
		LPCTSTR szInputBase64String);

	void InitBinary(
		LPSBinary lpsInputBin);

private:
	ULONG HandleChange(UINT nID);
	void UpdateParser();
};