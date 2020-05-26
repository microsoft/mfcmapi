#include <StdAfx.h>
#include <UI/file/exporter.h>
#include <UI/Dialogs/Editors/Editor.h>

#include <core/mapi/processor/dumpStore.h>
#include <core/mapi/extraPropTags.h>
#include <core/mapi/mapiMime.h>
#include <core/mapi/mapiFile.h>
#include <UI/mapiui.h> // TODO: Migrate export stuff from here

namespace exporter
{
	exportOptions getExportOptions(exporter::exportType exportType)
	{
		switch (exportType)
		{
		case exporter::exportType::text:
			return {L"xml", L".xml", strings::loadstring(IDS_XMLFILES)};
		case exporter::exportType::msgAnsi:
		case exporter::exportType::msgUnicode:
			return {L"msg", L".msg", strings::loadstring(IDS_MSGFILES)};
		case exporter::exportType::eml:
		case exporter::exportType::emlIConverter:
			return {L"eml", L".eml", strings::loadstring(IDS_EMLFILES)};
		case exporter::exportType::tnef:
			return {L"tnef", L".tnef", strings::loadstring(IDS_TNEFFILES)};
		default:
			return {};
		}
	}

	HRESULT exportMessage(
		exporter::exportType exportType,
		std::wstring filename,
		LPMESSAGE lpMessage,
		HWND hWnd,
		LPADRBOOK lpAddrBook)
	{
		auto hRes = S_OK;
		switch (exportType)
		{
		case exporter::exportType::text:
			// Idea is to capture anything that may be important about this message to disk so it can be analyzed.
			{
				mapi::processor::dumpStore MyDumpStore;
				MyDumpStore.InitMessagePath(filename);
				// Just assume this message might have attachments
				MyDumpStore.ProcessMessage(lpMessage, true, nullptr);
			}

			break;
		case exporter::exportType::msgAnsi:
			hRes = EC_H(file::SaveToMSG(lpMessage, filename, false, hWnd, true));
			break;
		case exporter::exportType::msgUnicode:
			hRes = EC_H(file::SaveToMSG(lpMessage, filename, true, hWnd, true));
			break;
		case exporter::exportType::eml:
			hRes = EC_H(file::SaveToEML(lpMessage, filename));
			break;
		case exporter::exportType::emlIConverter:
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
				hRes = EC_H(mapi::mapimime::ExportIMessageToEML(
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
		case exporter::exportType::tnef:
			hRes = EC_H(file::SaveToTNEF(lpMessage, lpAddrBook, filename));
			break;
		default:
			break;
		}

		return hRes;
	}
} // namespace exporter