#include <StdAfx.h>
#include <UI/file/exporter.h>
#include <UI/Dialogs/Editors/Editor.h>

#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <core/mapi/mapiABFunctions.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <core/mapi/columnTags.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <core/mapi/processor/dumpStore.h>
#include <core/utility/file.h>
#include <UI/Dialogs/ContentsTable/AttachmentsDlg.h>
#include <UI/MAPIFormFunctions.h>
#include <UI/file/FileDialogEx.h>
#include <core/mapi/extraPropTags.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <core/mapi/mapiProgress.h>
#include <core/mapi/mapiMime.h>
#include <core/sortlistdata/contentsData.h>
#include <core/mapi/cache/globalCache.h>
#include <core/mapi/mapiMemory.h>
#include <UI/mapiui.h>
#include <UI/addinui.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFile.h>
#include <core/interpret/flags.h>
#include <core/mapi/mapiFunctions.h>

namespace exporter
{
	//std::wstring getFileName(
	//	LPMESSAGE lpMessage,
	//	bool bPrompt,
	//	std::wstring szDotExt,
	//	std::wstring szExt,
	//	std::wstring dir,
	//	std::wstring szFilter)
	//{
	//	auto filename = file::BuildFileName(szDotExt, dir, lpMessage);
	//	output::DebugPrint(output::dbgLevel::Generic, L"BuildFileName built file name \"%ws\"\n", filename.c_str());

	//	if (bPrompt)
	//	{
	//		//filename =
	//		//	file::CFileDialogExW::SaveAs(szExt, filename, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, parent);
	//		filename =
	//			file::CFileDialogExW::SaveAs(szExt, filename, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, nullptr);
	//	}

	//	return filename;
	//}

	//void exportMessage(
	//	LPMESSAGE lpMessage,
	//	exportType iType,
	//	dialog::CBaseDialog * parent,
	//	LPADRBOOK lpAdrBook,
	//	bool bPrompt,
	//	std::wstring szDotExt,
	//	std::wstring szExt,
	//	std::wstring dir,
	//	std::wstring szFilter)
	//{
	//	auto hRes = S_OK;
	//	auto filename = file::BuildFileName(szDotExt, dir, lpMessage);
	//	output::DebugPrint(output::dbgLevel::Generic, L"BuildFileName built file name \"%ws\"\n", filename.c_str());

	//	if (bPrompt)
	//	{
	//		filename =
	//			file::CFileDialogExW::SaveAs(szExt, filename, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, parent);
	//	}

	//	if (!filename.empty())
	//	{
	//		switch (iType)
	//		{
	//		case exportType::text:
	//			// Idea is to capture anything that may be important about this message to disk so it can be analyzed.
	//			{
	//				mapi::processor::dumpStore MyDumpStore;
	//				MyDumpStore.InitMessagePath(filename);
	//				// Just assume this message might have attachments
	//				MyDumpStore.ProcessMessage(lpMessage, true, nullptr);
	//			}

	//			break;
	//		case exportType::msgAnsi:
	//			hRes = EC_H(file::SaveToMSG(lpMessage, filename, false, parent->m_hWnd, true));
	//			break;
	//		case exportType::msgUnicode:
	//			hRes = EC_H(file::SaveToMSG(lpMessage, filename, true, parent->m_hWnd, true));
	//			break;
	//		case exportType::eml:
	//			hRes = EC_H(file::SaveToEML(lpMessage, filename));
	//			break;
	//		case exportType::emlIConverter:
	//		{
	//			auto ulConvertFlags = CCSF_SMTP;
	//			auto et = IET_UNKNOWN;
	//			auto mst = USE_DEFAULT_SAVETYPE;
	//			ULONG ulWrapLines = USE_DEFAULT_WRAPPING;
	//			auto bDoAdrBook = false;

	//			hRes = EC_H(ui::mapiui::GetConversionToEMLOptions(
	//				parent, &ulConvertFlags, &et, &mst, &ulWrapLines, &bDoAdrBook));
	//			if (hRes == S_OK)
	//			{
	//				hRes = EC_H(mapi::mapimime::ExportIMessageToEML(
	//					lpMessage,
	//					filename.c_str(),
	//					ulConvertFlags,
	//					et,
	//					mst,
	//					ulWrapLines,
	//					bDoAdrBook ? lpAdrBook : nullptr));
	//			}
	//		}

	//		break;
	//		case exportType::tnef:
	//			hRes = EC_H(file::SaveToTNEF(lpMessage, lpAdrBook, filename));
	//			break;
	//		default:
	//			break;
	//		}
	//	}
	//	else
	//	{
	//		hRes = MAPI_E_USER_CANCEL;
	//	}

	//	lpMessage->Release();
	//}

	////needs:
	//// this - parent dialog - can provide
	//// hwnd - can provide or come from parent dialog
	//// selected item count
	//// address book
	//// we loop over selected items
	//void saveMessageToFile(HWND hwnd)
	//{
	//	auto hRes = S_OK;

	//	output::DebugPrint(output::dbgLevel::Generic, L"saveMessageToFile", L"\n");

	//	dialog::editor::CEditor MyData(
	//		this, IDS_SAVEMESSAGETOFILE, IDS_SAVEMESSAGETOFILEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	//	UINT uidDropDown[] = {IDS_DDTEXTFILE,
	//						  IDS_DDMSGFILEANSI,
	//						  IDS_DDMSGFILEUNICODE,
	//						  IDS_DDEMLFILE,
	//						  IDS_DDEMLFILEUSINGICONVERTERSESSION,
	//						  IDS_DDTNEFFILE};
	//	MyData.AddPane(
	//		viewpane::DropDownPane::Create(0, IDS_FORMATTOSAVEMESSAGE, _countof(uidDropDown), uidDropDown, true));
	//	const auto numSelected = m_lpContentsTableListCtrl->GetSelectedCount();
	//	if (numSelected > 1)
	//	{
	//		MyData.AddPane(viewpane::CheckPane::Create(1, IDS_EXPORTPROMPTLOCATION, false, false));
	//	}

	//	if (!MyData.DisplayDialog()) return;

	//	LPCWSTR szExt = nullptr;
	//	LPCWSTR szDotExt = nullptr;
	//	std::wstring szFilter;
	//	LPADRBOOK lpAddrBook = nullptr;
	//	switch (MyData.GetDropDown(0))
	//	{
	//	case 0:
	//		szExt = L"xml"; // STRING_OK
	//		szDotExt = L".xml"; // STRING_OK
	//		szFilter = strings::loadstring(IDS_XMLFILES);
	//		break;
	//	case 1:
	//	case 2:
	//		szExt = L"msg"; // STRING_OK
	//		szDotExt = L".msg"; // STRING_OK
	//		szFilter = strings::loadstring(IDS_MSGFILES);
	//		break;
	//	case 3:
	//	case 4:
	//		szExt = L"eml"; // STRING_OK
	//		szDotExt = L".eml"; // STRING_OK
	//		szFilter = strings::loadstring(IDS_EMLFILES);
	//		break;
	//	case 5:
	//		szExt = L"tnef"; // STRING_OK
	//		szDotExt = L".tnef"; // STRING_OK
	//		szFilter = strings::loadstring(IDS_TNEFFILES);

	//		lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
	//		break;
	//	default:
	//		break;
	//	}

	//	std::wstring dir;
	//	const auto bPrompt = numSelected == 1 || MyData.GetCheck(1);
	//	if (!bPrompt)
	//	{
	//		// If we weren't asked to prompt for each item, we still need to ask for a directory
	//		dir = file::GetDirectoryPath(hwnd);
	//	}

	//	auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);
	//	while (-1 != iItem)
	//	{
	//		auto lpMessage = OpenMessage(iItem, modifyType::REQUEST_MODIFY);
	//		if (lpMessage)
	//		{
	//			auto filename = file::BuildFileName(szDotExt, dir, lpMessage);
	//			output::DebugPrint(
	//				output::dbgLevel::Generic, L"BuildFileName built file name \"%ws\"\n", filename.c_str());

	//			if (bPrompt)
	//			{
	//				filename = file::CFileDialogExW::SaveAs(
	//					szExt, filename, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, this);
	//			}

	//			if (!filename.empty())
	//			{
	//				switch (MyData.GetDropDown(0))
	//				{
	//				case 0:
	//					// Idea is to capture anything that may be important about this message to disk so it can be analyzed.
	//					{
	//						mapi::processor::dumpStore MyDumpStore;
	//						MyDumpStore.InitMessagePath(filename);
	//						// Just assume this message might have attachments
	//						MyDumpStore.ProcessMessage(lpMessage, true, nullptr);
	//					}

	//					break;
	//				case 1:
	//					hRes = EC_H(file::SaveToMSG(lpMessage, filename, false, hwnd, true));
	//					break;
	//				case 2:
	//					hRes = EC_H(file::SaveToMSG(lpMessage, filename, true, hwnd, true));
	//					break;
	//				case 3:
	//					hRes = EC_H(file::SaveToEML(lpMessage, filename));
	//					break;
	//				case 4:
	//				{
	//					auto ulConvertFlags = CCSF_SMTP;
	//					auto et = IET_UNKNOWN;
	//					auto mst = USE_DEFAULT_SAVETYPE;
	//					ULONG ulWrapLines = USE_DEFAULT_WRAPPING;
	//					auto bDoAdrBook = false;

	//					hRes = EC_H(ui::mapiui::GetConversionToEMLOptions(
	//						this, &ulConvertFlags, &et, &mst, &ulWrapLines, &bDoAdrBook));
	//					if (hRes == S_OK)
	//					{
	//						LPADRBOOK lpAdrBook = nullptr;
	//						if (bDoAdrBook) lpAdrBook = m_lpMapiObjects->GetAddrBook(true); // do not release

	//						hRes = EC_H(mapi::mapimime::ExportIMessageToEML(
	//							lpMessage, filename.c_str(), ulConvertFlags, et, mst, ulWrapLines, lpAdrBook));
	//					}
	//				}

	//				break;
	//				case 5:
	//					hRes = EC_H(file::SaveToTNEF(lpMessage, lpAddrBook, filename));
	//					break;
	//				default:
	//					break;
	//				}
	//			}
	//			else
	//			{
	//				hRes = MAPI_E_USER_CANCEL;
	//			}

	//			lpMessage->Release();
	//		}
	//		else
	//		{
	//			hRes = MAPI_E_USER_CANCEL;
	//			CHECKHRESMSG(hRes, IDS_OPENMSGFAILED);
	//		}

	//		iItem = m_lpContentsTableListCtrl->GetNextItem(iItem, LVNI_SELECTED);
	//		if (hRes != S_OK && -1 != iItem)
	//		{
	//			if (bShouldCancel(this, hRes)) break;
	//		}
	//	}
	//}
} // namespace export