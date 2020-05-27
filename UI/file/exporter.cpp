#include <StdAfx.h>
#include <UI/file/exporter.h>
#include <UI/Dialogs/Editors/Editor.h>

#include <core/mapi/processor/dumpStore.h>
#include <core/mapi/extraPropTags.h>
#include <core/mapi/mapiMime.h>
#include <core/mapi/mapiFile.h>
#include <core/utility/file.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <UI/file/FileDialogEx.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/mapiui.h> // TODO: Migrate export stuff from here

namespace file
{
	bool exporter::init(
		CWnd* pParentWnd,
		bool bMultiSelect,
		HWND _hWnd,
		LPADRBOOK _lpAddrBook)
	{
		hWnd = _hWnd;
		lpAddrBook = _lpAddrBook;

		dialog::editor::CEditor MyData(
			pParentWnd, IDS_SAVEMESSAGETOFILE, IDS_SAVEMESSAGETOFILEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		UINT uidDropDown[] = {IDS_DDTEXTFILE,
							  IDS_DDMSGFILEANSI,
							  IDS_DDMSGFILEUNICODE,
							  IDS_DDEMLFILE,
							  IDS_DDEMLFILEUSINGICONVERTERSESSION,
							  IDS_DDTNEFFILE};
		MyData.AddPane(
			viewpane::DropDownPane::Create(0, IDS_FORMATTOSAVEMESSAGE, _countof(uidDropDown), uidDropDown, true));
		if (bMultiSelect)
		{
			MyData.AddPane(viewpane::CheckPane::Create(1, IDS_EXPORTPROMPTLOCATION, false, false));
		}

		if (!MyData.DisplayDialog()) return false;

		exportType = static_cast<file::exportType>(MyData.GetDropDown(0));
		bPrompt = !bMultiSelect || MyData.GetCheck(1);

		switch (exportType)
		{
		case exportType::text:
			szExt = L"xml";
			szDotExt = L".xml";
			szFilter = strings::loadstring(IDS_XMLFILES);
			break;
		case exportType::msgAnsi:
		case exportType::msgUnicode:
			szExt = L"msg";
			szDotExt = L".msg";
			szFilter = strings::loadstring(IDS_MSGFILES);
			break;
		case exportType::eml:
		case exportType::emlIConverter:
			szExt = L"eml";
			szDotExt = L".eml";
			szFilter = strings::loadstring(IDS_EMLFILES);
			break;
		case exportType::tnef:
			szExt = L"tnef";
			szDotExt = L".tnef";
			szFilter = strings::loadstring(IDS_TNEFFILES);
			break;
		}

		if (!bPrompt)
		{
			// If we weren't asked to prompt for each item, we still need to ask for a directory
			dir = file::GetDirectoryPath(hWnd);
		}

		return true;
	}

	HRESULT exporter::exportMessage(LPMESSAGE lpMessage)
	{
		auto hRes = S_OK;

		auto filename = file::BuildFileName(szDotExt, dir, lpMessage);
		output::DebugPrint(output::dbgLevel::Generic, L"BuildFileName built file name \"%ws\"\n", filename.c_str());

		if (bPrompt)
		{
			filename = file::CFileDialogExW::SaveAs(
				szExt, filename, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, CWnd::FromHandle(hWnd));
		}

		if (filename.empty()) return MAPI_E_USER_CANCEL;

		switch (exportType)
		{
		case exportType::text:
			// Idea is to capture anything that may be important about this message to disk so it can be analyzed.
			{
				mapi::processor::dumpStore MyDumpStore;
				MyDumpStore.InitMessagePath(filename);
				// Just assume this message might have attachments
				MyDumpStore.ProcessMessage(lpMessage, true, nullptr);
			}

			break;
		case exportType::msgAnsi:
			return EC_H(file::SaveToMSG(lpMessage, filename, false, hWnd, true));
			break;
		case exportType::msgUnicode:
			return EC_H(file::SaveToMSG(lpMessage, filename, true, hWnd, true));
			break;
		case exportType::eml:
			return EC_H(file::SaveToEML(lpMessage, filename));
			break;
		case exportType::emlIConverter:
		{
			auto ulConvertFlags = CCSF_SMTP;
			auto et = IET_UNKNOWN;
			auto mst = USE_DEFAULT_SAVETYPE;
			ULONG ulWrapLines = USE_DEFAULT_WRAPPING;
			auto bDoAdrBook = false;

			hRes = EC_H(ui::mapiui::GetConversionToEMLOptions(
				CWnd::FromHandle(hWnd), &ulConvertFlags, &et, &mst, &ulWrapLines, &bDoAdrBook));
			if (hRes == S_OK)
			{
				return EC_H(mapi::mapimime::ExportIMessageToEML(
					lpMessage,
					filename.c_str(),
					ulConvertFlags,
					et,
					mst,
					ulWrapLines,
					bDoAdrBook ? lpAddrBook : nullptr));
			}
		}

		break;
		case exportType::tnef:
			return EC_H(file::SaveToTNEF(lpMessage, lpAddrBook, filename));
			break;
		default:
			break;
		}

		return hRes;
	}
} // namespace file