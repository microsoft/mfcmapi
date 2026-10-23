#include <StdAfx.h>
#include <UI/Dialogs/Editors/MailboxEditor.h>
#include <UI/mapiui.h>
#include <core/mapi/mapiABFunctions.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <core/interpret/flags.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>

namespace dialog::editor
{
	CMailboxEditor::CMailboxEditor(_In_ CWnd* parent, _In_ LPMAPISESSION session)
		: CEditor(
			  parent,
			  IDS_OPENOTHERUSER,
			  IDS_OPENMAILBOXPROMPT,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2,
			  IDS_MAILBOXSELECTGAL,
			  IDS_MAILBOXADVANCED,
			  0),
		  m_session(session)
	{
		const UINT methods[] = {IDS_MAILBOXSMTP, IDS_MAILBOXGAL, IDS_MAILBOXDN, IDS_MAILBOXENTRYID};
		auto method = viewpane::DropDownPane::Create(0, IDS_MAILBOXMETHOD, _countof(methods), methods, true);
		method->SetDropDownSelection(0);
		AddPane(method);
		AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_MAILBOXSMTP, false));
		m_options[static_cast<size_t>(MailboxMethod::GAL)].flags =
			OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP;
		m_options[static_cast<size_t>(MailboxMethod::DN)].flags =
			OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP;
		m_options[static_cast<size_t>(MailboxMethod::EntryID)].flags = MDB_WRITE;
	}

	BOOL CMailboxEditor::OnInitDialog()
	{
		const auto result = CEditor::OnInitDialog();
		UpdateMethod();
		return result;
	}

	void CMailboxEditor::UpdateMethod()
	{
		const UINT labels[] = {IDS_MAILBOXSMTP, IDS_MAILBOXGALIDENTITY, IDS_USERDN, IDS_EID};
		GetPane(1)->SetLabel(labels[static_cast<size_t>(m_method)]);
		SetStringW(1, Options().identity);
		GetDlgItem(IDD_EDITACTION1)->EnableWindow(m_method == MailboxMethod::GAL);
		GetDlgItem(IDD_EDITACTION2)->EnableWindow(m_method != MailboxMethod::SMTP);
		OnRecalcLayout();
	}

	ULONG CMailboxEditor::HandleChange(UINT nID)
	{
		const auto pane = CEditor::HandleChange(nID);
		if (pane == 0)
		{
			const auto method = GetDropDown(0);
			if (method >= 0 && method < static_cast<int>(m_options.size()))
			{
				Options().identity = GetStringW(1);
				m_method = static_cast<MailboxMethod>(method);
				UpdateMethod();
			}
		}
		return pane;
	}

	void CMailboxEditor::OnEditAction1()
	{
		if (m_method != MailboxMethod::GAL) return;
		LPADRBOOK addressBook{};
		const auto openResult = EC_MAPI(
			m_session->OpenAddressBook(reinterpret_cast<ULONG_PTR>(m_hWnd), nullptr, AB_NO_DIALOG, &addressBook));
		if (FAILED(openResult) || !addressBook)
		{
			if (addressBook) addressBook->Release();
			if (SUCCEEDED(openResult)) EC_H_S(MAPI_E_INVALID_OBJECT);
			return;
		}
		auto user = mapi::ab::SelectUser(addressBook, m_hWnd, nullptr);
		addressBook->Release();
		if (!user) return;

		static const SizedSPropTagArray(1, tags) = {1, {PR_EMAIL_ADDRESS_W}};
		ULONG count{};
		LPSPropValue props{};
		const auto hRes = EC_MAPI(user->GetProps(LPSPropTagArray(&tags), MAPI_UNICODE, &count, &props));
		if (SUCCEEDED(hRes))
		{
			const auto dn = props ? PpropFindProp(props, count, PR_EMAIL_ADDRESS_W) : nullptr;
			if (strings::CheckStringProp(dn, PT_UNICODE) && dn->Value.lpszW[0])
			{
				Options().identity = dn->Value.lpszW;
				SetStringW(1, Options().identity);
			}
			else
			{
				EC_H_S(MAPI_E_NOT_FOUND);
			}
		}
		MAPIFreeBuffer(props);
		user->Release();
	}

	void CMailboxEditor::OnEditAction2()
	{
		if (m_method == MailboxMethod::SMTP) return;
		const auto entryID = m_method == MailboxMethod::EntryID;
		auto& options = Options();
		CEditor advanced(
			this,
			IDS_MAILBOXADVANCED,
			entryID ? IDS_OPENSTOREEIDPROMPT : IDS_OPENWITHFLAGSPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		advanced.AddPane(
			viewpane::TextPane::CreateSingleLinePane(0, entryID ? IDS_FLAGS : IDS_CREATESTORENTRYIDFLAGS, false));
		advanced.SetHex(0, options.flags);
		if (entryID)
		{
			advanced.AddPane(viewpane::CheckPane::Create(1, IDS_EIDBASE64ENCODED, options.base64, false));
			advanced.AddPane(viewpane::CheckPane::Create(2, IDS_DISPLAYPROPS, options.displayProps, false));
			advanced.AddPane(viewpane::CheckPane::Create(3, IDS_UNWRAPSTORE, options.unwrap, false));
		}
		else
		{
			advanced.SetPromptPostFix(flags::AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
			const auto server = options.server.empty() ? mapi::store::GetServerName(m_session) : options.server;
			advanced.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_SERVERNAME, server, false));
			advanced.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_USER_SMTP_ADDRESS, options.smtp, false));
			advanced.AddPane(viewpane::CheckPane::Create(3, IDS_FORCESERVER, options.forceServer, false));
		}
		if (!advanced.DisplayDialog()) return;
		options.flags = advanced.GetHex(0);
		if (entryID)
		{
			options.base64 = advanced.GetCheck(1);
			options.displayProps = advanced.GetCheck(2);
			options.unwrap = advanced.GetCheck(3);
		}
		else
		{
			options.server = advanced.GetStringW(1);
			options.smtp = advanced.GetStringW(2);
			options.forceServer = advanced.GetCheck(3);
		}
	}

	void CMailboxEditor::OnOK()
	{
		Options().identity = GetStringW(1);
		if (m_method == MailboxMethod::SMTP && !mapi::ab::IsMailboxSMTPAddress(Options().identity))
		{
			ui::mapiui::displayError(strings::loadstring(IDS_MAILBOXINVALIDSMTP));
			return;
		}
		if (Options().identity.empty() ||
			(m_method == MailboxMethod::EntryID && GetBinary(1, Options().base64).empty()))
		{
			ui::mapiui::displayError(strings::loadstring(IDS_MAILBOXINVALIDIDENTITY));
			return;
		}
		CEditor::OnOK();
	}
} // namespace dialog::editor
