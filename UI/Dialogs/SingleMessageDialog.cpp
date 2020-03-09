#include <StdAfx.h>
#include <UI/Dialogs/SingleMessageDialog.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/StreamEditor.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"SingleMessageDialog";

	SingleMessageDialog::SingleMessageDialog(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_opt_ LPMAPIPROP lpMAPIProp)
		: CBaseDialog(pParentWnd, lpMapiObjects, NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_szTitle = strings::loadstring(IDS_MESSAGE);

		m_lpMessage = mapi::safe_cast<LPMESSAGE>(lpMAPIProp);

		CBaseDialog::CreateDialogAndMenu(IDR_MENU_MESSAGE, NULL, NULL);
	}

	SingleMessageDialog::~SingleMessageDialog()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpMessage) m_lpMessage->Release();
		m_lpMessage = nullptr;
	}

	BOOL SingleMessageDialog::OnInitDialog()
	{
		const auto bRet = CBaseDialog::OnInitDialog();

		if (m_lpMessage)
		{
			// Get a property for the title bar
			m_szTitle = mapi::GetTitle(m_lpMessage);
		}

		UpdateTitleBarText();

		EC_H_S(m_lpPropDisplay->SetDataSource(m_lpMessage, NULL, false));

		return bRet;
	}

	BEGIN_MESSAGE_MAP(SingleMessageDialog, CBaseDialog)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	ON_COMMAND(ID_ATTACHMENTPROPERTIES, OnAttachmentProperties)
	ON_COMMAND(ID_RECIPIENTPROPERTIES, OnRecipientProperties)
	ON_COMMAND(ID_RTFSYNC, OnRTFSync)
	ON_COMMAND(ID_TESTEDITBODY, OnTestEditBody)
	ON_COMMAND(ID_TESTEDITHTML, OnTestEditHTML)
	ON_COMMAND(ID_TESTEDITRTF, OnTestEditRTF)
	ON_COMMAND(ID_SAVECHANGES, OnSaveChanges)
	END_MESSAGE_MAP()

	// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
	void SingleMessageDialog::OnRefreshView()
	{
		if (!m_lpPropDisplay) return;
		static_cast<void>(m_lpPropDisplay->RefreshMAPIPropList());
	}

	void SingleMessageDialog::OnAttachmentProperties()
	{
		if (!m_lpMessage) return;

		EC_H_S(DisplayTable(m_lpMessage, PR_MESSAGE_ATTACHMENTS, objectType::default, this));
	}

	void SingleMessageDialog::OnRecipientProperties()
	{
		if (!m_lpMessage) return;

		EC_H_S(DisplayTable(m_lpMessage, PR_MESSAGE_RECIPIENTS, objectType::default, this));
	}

	void SingleMessageDialog::OnRTFSync()
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		editor::CEditor MyData(this, IDS_CALLRTFSYNC, IDS_CALLRTFSYNCPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FLAGS, false));
		MyData.SetHex(0, RTF_SYNC_RTF_CHANGED);

		if (MyData.DisplayDialog())
		{
			BOOL bMessageUpdated = false;

			if (m_lpMessage)
			{
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"Calling RTFSync on %p with flags 0x%X\n",
					m_lpMessage,
					MyData.GetHex(0));
				hRes = EC_MAPI(RTFSync(m_lpMessage, MyData.GetHex(0), &bMessageUpdated));
				output::DebugPrint(output::dbgLevel::Generic, L"RTFSync returned %d\n", bMessageUpdated);

				if (SUCCEEDED(hRes))
				{
					EC_MAPI_S(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
				}

				static_cast<void>(m_lpPropDisplay->RefreshMAPIPropList());
			}
		}
	}

	void SingleMessageDialog::OnTestEditBody()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (m_lpMessage)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"Editing body on %p\n", m_lpMessage);

			editor::CStreamEditor MyEditor(
				this,
				IDS_PROPEDITOR,
				IDS_STREAMEDITORPROMPT,
				m_lpMessage,
				PR_BODY_W,
				false, // No stream guessing
				m_bIsAB,
				false,
				false,
				0,
				0,
				0);
			MyEditor.DisableSave();

			static_cast<void>(MyEditor.DisplayDialog());

			static_cast<void>(m_lpPropDisplay->RefreshMAPIPropList());
		}
	}

	void SingleMessageDialog::OnTestEditHTML()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (m_lpMessage)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"Testing HTML on %p\n", m_lpMessage);

			editor::CStreamEditor MyEditor(
				this,
				IDS_PROPEDITOR,
				IDS_STREAMEDITORPROMPT,
				m_lpMessage,
				PR_BODY_HTML,
				false, // No stream guessing
				m_bIsAB,
				false,
				false,
				0,
				0,
				0);
			MyEditor.DisableSave();

			static_cast<void>(MyEditor.DisplayDialog());

			static_cast<void>(m_lpPropDisplay->RefreshMAPIPropList());
		}
	}

	void SingleMessageDialog::OnTestEditRTF()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (m_lpMessage)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"Testing body on %p\n", m_lpMessage);

			editor::CStreamEditor MyEditor(
				this,
				IDS_PROPEDITOR,
				IDS_STREAMEDITORPROMPT,
				m_lpMessage,
				PR_RTF_COMPRESSED,
				false, // No stream guessing
				m_bIsAB,
				true,
				false,
				0,
				0,
				0);
			MyEditor.DisableSave();

			static_cast<void>(MyEditor.DisplayDialog());

			static_cast<void>(m_lpPropDisplay->RefreshMAPIPropList());
		}
	}

	void SingleMessageDialog::OnSaveChanges()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (m_lpMessage)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"Saving changes on %p\n", m_lpMessage);

			EC_MAPI_S(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

			static_cast<void>(m_lpPropDisplay->RefreshMAPIPropList());
		}
	}
} // namespace dialog