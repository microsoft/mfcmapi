#include <StdAfx.h>
#include <UI/mapiui.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <Interpret/Guids.h>
#include <MAPI/MAPIProgress.h>
#include <MAPI/MAPIFunctions.h>
#include <IO/File.h>
#include <UI/FileDialogEx.h>

namespace ui
{
	namespace mapiui
	{
		void initCallbacks()
		{
			mapi::GetCopyDetails = [](auto _1, auto _2, auto _3, auto _4, auto _5) -> mapi::CopyDetails {
				return GetCopyDetails(_1, _2, _3, _4, _5);
			};
		}

		// Takes a tag array (and optional MAPIProp) and displays UI prompting to build an exclusion array
		// Must be freed with MAPIFreeBuffer
		LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB)
		{
			dialog::editor::CTagArrayEditor TagEditor(
				nullptr, IDS_TAGSTOEXCLUDE, IDS_TAGSTOEXCLUDEPROMPT, nullptr, lpTagArray, bIsAB, lpProp);
			if (TagEditor.DisplayDialog())
			{
				return TagEditor.DetachModifiedTagArray();
			}

			return nullptr;
		}

		mapi::CopyDetails GetCopyDetails(
			HWND hWnd,
			_In_ LPMAPIPROP lpSource,
			LPCGUID lpGUID,
			_In_opt_ LPSPropTagArray lpTagArray,
			bool bIsAB)
		{
			dialog::editor::CEditor MyData(
				nullptr, IDS_COPYTO, IDS_COPYPASTEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.AddPane(
				viewpane::TextPane::CreateSingleLinePane(0, IDS_INTERFACE, guid::GUIDToStringAndName(lpGUID), false));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_FLAGS, false));
			MyData.SetHex(1, MAPI_DIALOG);

			if (!MyData.DisplayDialog()) return {};

			const auto progress = hWnd ? mapi::mapiui::GetMAPIProgress(L"CopyTo", hWnd) : LPMAPIPROGRESS{};

			return {true,
					MyData.GetHex(1) | (progress ? MAPI_DIALOG : 0),
					guid::StringToGUID(MyData.GetStringW(0)),
					progress,
					progress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
					GetExcludedTags(lpTagArray, lpSource, bIsAB),
					true};
		}

		void ExportMessages(_In_ LPMAPIFOLDER lpFolder, HWND hWnd)
		{
			dialog::editor::CEditor MyData(
				nullptr, IDS_EXPORTTITLE, IDS_EXPORTPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_EXPORTSEARCHTERM, false));

			if (!MyData.DisplayDialog()) return;

			auto szDir = file::GetDirectoryPath(hWnd);
			if (szDir.empty()) return;
			auto restrictString = MyData.GetStringW(0);
			if (restrictString.empty()) return;

			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			LPMAPITABLE lpTable = nullptr;
			WC_MAPI_S(lpFolder->GetContentsTable(MAPI_DEFERRED_ERRORS | MAPI_UNICODE, &lpTable));
			if (lpTable)
			{
				// Allocate and create our SRestriction
				auto lpRes = mapi::CreatePropertyStringRestriction(
					PR_SUBJECT_W, restrictString, FL_SUBSTRING | FL_IGNORECASE, nullptr);
				if (lpRes)
				{
					auto hRes = WC_MAPI(lpTable->Restrict(lpRes, 0));
					MAPIFreeBuffer(lpRes);
					if (SUCCEEDED(hRes))
					{
						enum
						{
							fldPR_ENTRYID,
							fldPR_SUBJECT_W,
							fldPR_RECORD_KEY,
							fldNUM_COLS
						};

						static const SizedSPropTagArray(fldNUM_COLS, fldCols) = {
							fldNUM_COLS, {PR_ENTRYID, PR_SUBJECT_W, PR_RECORD_KEY}};

						hRes = WC_MAPI(lpTable->SetColumns(LPSPropTagArray(&fldCols), TBL_ASYNC));

						// Export messages in the rows
						LPSRowSet lpRows = nullptr;
						if (SUCCEEDED(hRes))
						{
							for (;;)
							{
								if (lpRows) FreeProws(lpRows);
								lpRows = nullptr;
								WC_MAPI_S(lpTable->QueryRows(50, NULL, &lpRows));
								if (!lpRows || !lpRows->cRows) break;

								for (ULONG i = 0; i < lpRows->cRows; i++)
								{
									WC_H_S(file::SaveToMSG(
										lpFolder,
										szDir,
										lpRows->aRow[i].lpProps[fldPR_ENTRYID],
										&lpRows->aRow[i].lpProps[fldPR_RECORD_KEY],
										&lpRows->aRow[i].lpProps[fldPR_SUBJECT_W],
										true,
										hWnd));
								}
							}

							if (lpRows) FreeProws(lpRows);
						}
					}
				}

				lpTable->Release();
			}
		}

		_Check_return_ HRESULT WriteAttachmentToFile(_In_ LPATTACH lpAttach, HWND hWnd)
		{
			if (!lpAttach) return MAPI_E_INVALID_PARAMETER;

			enum
			{
				ATTACH_METHOD,
				ATTACH_LONG_FILENAME_W,
				ATTACH_FILENAME_W,
				DISPLAY_NAME_W,
				NUM_COLS
			};
			static const SizedSPropTagArray(NUM_COLS, sptaAttachProps) = {
				NUM_COLS, {PR_ATTACH_METHOD, PR_ATTACH_LONG_FILENAME_W, PR_ATTACH_FILENAME_W, PR_DISPLAY_NAME_W}};

			output::DebugPrint(DBGGeneric, L"WriteAttachmentToFile: Saving attachment.\n");

			LPSPropValue lpProps = nullptr;
			ULONG ulProps = 0;

			// Get required properties from the message
			auto hRes = EC_H_GETPROPS(lpAttach->GetProps(
				LPSPropTagArray(&sptaAttachProps), // property tag array
				fMapiUnicode, // flags
				&ulProps, // Count of values returned
				&lpProps)); // Values returned
			if (lpProps)
			{
				auto szName = L"Unknown"; // STRING_OK

				// Get a file name to use
				if (mapi::CheckStringProp(&lpProps[ATTACH_LONG_FILENAME_W], PT_UNICODE))
				{
					szName = lpProps[ATTACH_LONG_FILENAME_W].Value.lpszW;
				}
				else if (mapi::CheckStringProp(&lpProps[ATTACH_FILENAME_W], PT_UNICODE))
				{
					szName = lpProps[ATTACH_FILENAME_W].Value.lpszW;
				}
				else if (mapi::CheckStringProp(&lpProps[DISPLAY_NAME_W], PT_UNICODE))
				{
					szName = lpProps[DISPLAY_NAME_W].Value.lpszW;
				}

				auto szFileName = strings::SanitizeFileName(szName);

				// Get File Name
				switch (lpProps[ATTACH_METHOD].Value.l)
				{
				case ATTACH_BY_VALUE:
				case ATTACH_BY_REFERENCE:
				case ATTACH_BY_REF_RESOLVE:
				case ATTACH_BY_REF_ONLY:
				{
					output::DebugPrint(
						DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());

					auto file = file::CFileDialogExW::SaveAs(
						L"txt", // STRING_OK
						szFileName,
						OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
						strings::loadstring(IDS_ALLFILES));
					if (!file.empty())
					{
						hRes = EC_H(file::WriteAttachStreamToFile(lpAttach, file));
					}
				}
				break;
				case ATTACH_EMBEDDED_MSG:
					// Get File Name
					{
						output::DebugPrint(
							DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());
						auto file = file::CFileDialogExW::SaveAs(
							L"msg", // STRING_OK
							szFileName,
							OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
							strings::loadstring(IDS_MSGFILES));
						if (!file.empty())
						{
							hRes =
								EC_H(file::WriteEmbeddedMSGToFile(lpAttach, file, MAPI_UNICODE == fMapiUnicode, hWnd));
						}
					}
					break;
				case ATTACH_OLE:
				{
					output::DebugPrint(
						DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());
					auto file = file::CFileDialogExW::SaveAs(
						strings::emptystring,
						szFileName,
						OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
						strings::loadstring(IDS_ALLFILES));
					if (!file.empty())
					{
						hRes = EC_H(file::WriteOleToFile(lpAttach, file));
					}
				}
				break;
				default:
					error::ErrDialog(__FILE__, __LINE__, IDS_EDUNKNOWNATTMETHOD);
					break;
				}
			}

			MAPIFreeBuffer(lpProps);
			return hRes;
		}
	} // namespace mapiui
} // namespace ui