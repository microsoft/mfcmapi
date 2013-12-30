// Options.cpp : implementation file
//

#include "stdafx.h"
#include "Editor.h"
#include "Options.h"
#include "ImportProcs.h"

class COptions : public CEditor
{
public:
	COptions(
		_In_ CWnd* pParentWnd);
	virtual ~COptions();
	bool NeedPropRefresh();

private:
	_Check_return_ ULONG HandleChange(UINT nID);

	void OnOK();

	bool m_bNeedPropRefresh;
};

static TCHAR* CLASS = _T("COptions");

COptions::COptions(_In_ CWnd* pWnd):
CEditor(pWnd, IDS_SETOPTS, NULL, 0, CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(CLASS);
	EnableScroll();

	CreateControls(NumRegOptionKeys);

	m_bNeedPropRefresh = false;

	DebugPrintEx(DBGGeneric, CLASS, _T("COptions("), _T("Building option sheet - adding fields\n"));

	ULONG ulReg = 0;

	for (ulReg = 0 ; ulReg < NumRegOptionKeys ; ulReg++)
	{
		if (regoptCheck == RegKeys[ulReg].ulRegOptType)
		{
			InitPane(ulReg, CreateCheckPane(RegKeys[ulReg].uiOptionsPrompt, (0 != RegKeys[ulReg].ulCurDWORD), false));
		}
		else if (regoptString == RegKeys[ulReg].ulRegOptType)
		{
			InitPane(ulReg, CreateSingleLinePane(RegKeys[ulReg].uiOptionsPrompt, RegKeys[ulReg].szCurSTRING, false));
		}
		else if (regoptStringHex == RegKeys[ulReg].ulRegOptType)
		{
			InitPane(ulReg, CreateSingleLinePane(RegKeys[ulReg].uiOptionsPrompt, NULL, false));
			SetHex(ulReg, RegKeys[ulReg].ulCurDWORD);
		}
		else if (regoptStringDec == RegKeys[ulReg].ulRegOptType)
		{
			InitPane(ulReg, CreateSingleLinePane(RegKeys[ulReg].uiOptionsPrompt, NULL, false));
			SetDecimal(ulReg, RegKeys[ulReg].ulCurDWORD);
		}
	}
} // COptions::CHexEditor

COptions::~COptions()
{
	TRACE_DESTRUCTOR(CLASS);
} // COptions::~CHexEditor

void COptions::OnOK()
{
	// Do OK work
	HRESULT hRes = S_OK;
	ULONG ulReg = 0;

	// need to grab this FIRST
	EC_H(StringCchCopy(RegKeys[regkeyDEBUG_FILE_NAME].szCurSTRING, _countof(RegKeys[regkeyDEBUG_FILE_NAME].szCurSTRING), GetStringUseControl(regkeyDEBUG_FILE_NAME)));

	if (GetHex(regkeyDEBUG_TAG) != RegKeys[regkeyDEBUG_TAG].ulCurDWORD)
	{
		SetDebugLevel(GetHex(regkeyDEBUG_TAG));
		DebugPrintVersion(DBGVersionBanner);
	}

	SetDebugOutputToFile(GetCheck(regkeyDEBUG_TO_FILE));

	// Remaining options require no special handling - loop through them
	for (ulReg = 0 ; ulReg < NumRegOptionKeys ; ulReg++)
	{
		if (regoptCheck == RegKeys[ulReg].ulRegOptType)
		{
			if (RegKeys[ulReg].bRefresh && RegKeys[ulReg].ulCurDWORD != (ULONG) GetCheckUseControl(ulReg))
			{
				m_bNeedPropRefresh = true;
			}
			RegKeys[ulReg].ulCurDWORD = GetCheckUseControl(ulReg);
		}
		else if (regoptStringHex == RegKeys[ulReg].ulRegOptType)
		{
			RegKeys[ulReg].ulCurDWORD = GetHexUseControl(ulReg);
		}
		else if (regoptStringDec == RegKeys[ulReg].ulRegOptType)
		{
			RegKeys[ulReg].ulCurDWORD = GetDecimalUseControl(ulReg);
		}
	}

	// Commit our values to the registry
	WriteToRegistry();

	CEditor::OnOK();
} // COptions::OnOK

_Check_return_ ULONG COptions::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);
	return i;
} // COptions::HandleChange

bool COptions::NeedPropRefresh()
{
	return m_bNeedPropRefresh;
} // COptions::NeedPropRefresh

bool DisplayOptionsDlg(_In_ CWnd* lpParentWnd)
{
	COptions MyOptions(lpParentWnd);
	MyOptions.DoModal();

	ForceOutlookMAPI(0 != RegKeys[regkeyFORCEOUTLOOKMAPI].ulCurDWORD);
	ForceSystemMAPI(0 != RegKeys[regkeyFORCESYSTEMMAPI].ulCurDWORD);
	return MyOptions.NeedPropRefresh();
} // DisplayOptionsDlg