#pragma once
// SmartViewPane.h : header file

#include "DropDownPane.h"
#include "TextPane.h"

ViewPane* CreateSmartViewPane(UINT uidLabel);

class SmartViewPane : public DropDownPane
{
public:
	SmartViewPane(UINT uidLabel);
	virtual ~SmartViewPane();

	virtual bool IsType(__ViewTypes vType);
	virtual ULONG GetFlags();
	virtual void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc);
	virtual void SetWindowPos(int x, int y, int width, int height);
	virtual int GetFixedHeight();
	virtual int GetLines();

	virtual void SetMargins(
		int iMargin,
		int iSideMargin,
		int iLabelHeight, // Height of the label
		int iSmallHeightMargin,
		int iLargeHeightMargin,
		int iButtonHeight, // Height of buttons below the control
		int iEditHeight); // height of an edit control

	void SetStringW(_In_opt_z_ LPCWSTR szMsg);
	void DisableDropDown();
	void SetParser(__ParsingTypeEnum iParser);

	void Parse(SBinary myBin);

private:
	TextPane* m_lpTextPane;

	bool m_bHasData;
	bool m_bDoDropDown;
};