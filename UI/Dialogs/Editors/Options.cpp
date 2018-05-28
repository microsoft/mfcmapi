#include <StdAfx.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/Options.h>
#include <ImportProcs.h>

namespace editor
{
	class COptions : public CEditor
	{
	public:
		COptions(_In_ CWnd* pParentWnd);
		virtual ~COptions();
		bool NeedPropRefresh() const;

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;

		void OnOK() override;

		bool m_bNeedPropRefresh;
	};

	static std::wstring CLASS = L"COptions";

	COptions::COptions(_In_ CWnd* pWnd) :
		CEditor(pWnd, IDS_SETOPTS, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		EnableScroll();

		m_bNeedPropRefresh = false;

		DebugPrintEx(DBGGeneric, CLASS, L"COptions(", L"Building option sheet - adding fields\n");

		for (ULONG ulReg = 0; ulReg < NumRegOptionKeys; ulReg++)
		{
			if (regoptCheck == RegKeys[ulReg].ulRegOptType)
			{
				InitPane(ulReg, viewpane::CheckPane::Create(RegKeys[ulReg].uiOptionsPrompt, 0 != RegKeys[ulReg].ulCurDWORD, false));
			}
			else if (regoptString == RegKeys[ulReg].ulRegOptType)
			{
				InitPane(ulReg, viewpane::TextPane::CreateSingleLinePane(RegKeys[ulReg].uiOptionsPrompt, RegKeys[ulReg].szCurSTRING, false));
			}
			else if (regoptStringHex == RegKeys[ulReg].ulRegOptType)
			{
				InitPane(ulReg, viewpane::TextPane::CreateSingleLinePane(RegKeys[ulReg].uiOptionsPrompt, false));
				SetHex(ulReg, RegKeys[ulReg].ulCurDWORD);
			}
			else if (regoptStringDec == RegKeys[ulReg].ulRegOptType)
			{
				InitPane(ulReg, viewpane::TextPane::CreateSingleLinePane(RegKeys[ulReg].uiOptionsPrompt, false));
				SetDecimal(ulReg, RegKeys[ulReg].ulCurDWORD);
			}
		}
	}

	COptions::~COptions()
	{
		TRACE_DESTRUCTOR(CLASS);
	}

	void COptions::OnOK()
	{
		// need to grab this FIRST
		RegKeys[regkeyDEBUG_FILE_NAME].szCurSTRING = GetStringW(regkeyDEBUG_FILE_NAME);

		if (GetHex(regkeyDEBUG_TAG) != RegKeys[regkeyDEBUG_TAG].ulCurDWORD)
		{
			SetDebugLevel(GetHex(regkeyDEBUG_TAG));
			DebugPrintVersion(DBGVersionBanner);
		}

		SetDebugOutputToFile(GetCheck(regkeyDEBUG_TO_FILE));

		// Remaining options require no special handling - loop through them
		for (ULONG ulReg = 0; ulReg < NumRegOptionKeys; ulReg++)
		{
			if (regoptCheck == RegKeys[ulReg].ulRegOptType)
			{
				if (RegKeys[ulReg].bRefresh && RegKeys[ulReg].ulCurDWORD != static_cast<ULONG>(GetCheck(ulReg)))
				{
					m_bNeedPropRefresh = true;
				}

				RegKeys[ulReg].ulCurDWORD = GetCheck(ulReg);
			}
			else if (regoptStringHex == RegKeys[ulReg].ulRegOptType)
			{
				RegKeys[ulReg].ulCurDWORD = GetHex(ulReg);
			}
			else if (regoptStringDec == RegKeys[ulReg].ulRegOptType)
			{
				RegKeys[ulReg].ulCurDWORD = GetDecimal(ulReg);
			}
		}

		// Commit our values to the registry
		WriteToRegistry();

		CEditor::OnOK();
	}

	_Check_return_ ULONG COptions::HandleChange(UINT nID)
	{
		return CEditor::HandleChange(nID);
	}

	bool COptions::NeedPropRefresh() const
	{
		return m_bNeedPropRefresh;
	}

	bool DisplayOptionsDlg(_In_ CWnd* lpParentWnd)
	{
		COptions MyOptions(lpParentWnd);
		MyOptions.DoModal();

		ForceOutlookMAPI(0 != RegKeys[regkeyFORCEOUTLOOKMAPI].ulCurDWORD);
		ForceSystemMAPI(0 != RegKeys[regkeyFORCESYSTEMMAPI].ulCurDWORD);
		return MyOptions.NeedPropRefresh();
	}
}