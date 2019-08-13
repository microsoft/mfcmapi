#include <StdAfx.h>
#include <UI/mapiui.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <core/interpret/guid.h>
#include <core/mapi/mapiProgress.h>
#include <core/utility/file.h>
#include <UI/FileDialogEx.h>
#include <core/mapi/extraPropTags.h>
#include <core/mapi/mapiMime.h>
#include <core/utility/import.h>
#include <UI/UIFunctions.h>
#include <core/mapi/mapiABFunctions.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <UI/CMAPIProgress.h>
#include <UI/Dialogs/Editors/DbgView.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFile.h>
#include <core/interpret/flags.h>
#include <core/mapi/mapiFunctions.h>
#include <UI/OnNotify.h>
#include <UI/AdviseSink.h>

namespace ui
{
	namespace mapiui
	{
		void initCallbacks()
		{
			mapi::GetCopyDetails = [](auto _1, auto _2, auto _3, auto _4, auto _5) -> mapi::CopyDetails {
				return GetCopyDetails(_1, _2, _3, _4, _5);
			};
			error::displayError = [](auto _1) { displayError(_1); };
			mapi::mapiui::getMAPIProgress = [](auto _1, auto _2) -> LPMAPIPROGRESS {
				return ui::GetMAPIProgress(_1, _2);
			};
			output::outputToDbgView = [](auto _1) { OutputToDbgView(_1); };
			mapi::store::promptServerName = []() { return PromptServerName(); };
			mapi::mapiui::onNotifyCallback = [](auto _1, auto _2, auto _3, auto _4) {
				mapi::mapiui::OnNotify(_1, _2, _3, _4);
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

			output::DebugPrint(output::DBGGeneric, L"WriteAttachmentToFile: Saving attachment.\n");

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
				if (strings::CheckStringProp(&lpProps[ATTACH_LONG_FILENAME_W], PT_UNICODE))
				{
					szName = lpProps[ATTACH_LONG_FILENAME_W].Value.lpszW;
				}
				else if (strings::CheckStringProp(&lpProps[ATTACH_FILENAME_W], PT_UNICODE))
				{
					szName = lpProps[ATTACH_FILENAME_W].Value.lpszW;
				}
				else if (strings::CheckStringProp(&lpProps[DISPLAY_NAME_W], PT_UNICODE))
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
						output::DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());

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
							output::DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());
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
						output::DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());
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

		_Check_return_ HRESULT GetConversionToEMLOptions(
			_In_ CWnd* pParentWnd,
			_Out_opt_ CCSFLAGS* lpConvertFlags,
			_Out_opt_ ENCODINGTYPE* lpet,
			_Out_opt_ MIMESAVETYPE* lpmst,
			_Out_opt_ ULONG* lpulWrapLines,
			_Out_opt_ bool* pDoAdrBook)
		{
			if (lpConvertFlags) *lpConvertFlags = {};
			if (lpet) *lpet = {};
			if (lpmst) *lpmst = {};
			if (lpulWrapLines) *lpulWrapLines = {};
			if (pDoAdrBook) *pDoAdrBook = {};
			if (!lpConvertFlags || !lpet || !lpmst || !lpulWrapLines || !pDoAdrBook) return MAPI_E_INVALID_PARAMETER;

			dialog::editor::CEditor MyData(
				pParentWnd, IDS_CONVERTTOEML, IDS_CONVERTTOEMLPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_CONVERTFLAGS, false));
			MyData.SetHex(0, CCSF_SMTP);
			MyData.AddPane(viewpane::CheckPane::Create(1, IDS_CONVERTDOENCODINGTYPE, false, false));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_CONVERTENCODINGTYPE, false));
			MyData.SetHex(2, IET_7BIT);
			MyData.AddPane(viewpane::CheckPane::Create(3, IDS_CONVERTDOMIMESAVETYPE, false, false));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_CONVERTMIMESAVETYPE, false));
			MyData.SetHex(4, SAVE_RFC822);
			MyData.AddPane(viewpane::CheckPane::Create(5, IDS_CONVERTDOWRAPLINES, false, false));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(6, IDS_CONVERTWRAPLINECOUNT, false));
			MyData.SetDecimal(6, 74);
			MyData.AddPane(viewpane::CheckPane::Create(7, IDS_CONVERTDOADRBOOK, false, false));

			if (!MyData.DisplayDialog()) return MAPI_E_USER_CANCEL;

			*lpConvertFlags = static_cast<CCSFLAGS>(MyData.GetHex(0));
			*lpet = MyData.GetCheck(1) ? static_cast<ENCODINGTYPE>(MyData.GetDecimal(2)) : IET_UNKNOWN;
			*lpmst = MyData.GetCheck(3) ? static_cast<MIMESAVETYPE>(MyData.GetHex(4)) : USE_DEFAULT_SAVETYPE;
			*lpulWrapLines = MyData.GetCheck(5) ? MyData.GetDecimal(6) : USE_DEFAULT_WRAPPING;
			*pDoAdrBook = MyData.GetCheck(7);

			return S_OK;
		}

		_Check_return_ HRESULT GetConversionFromEMLOptions(
			_In_ CWnd* pParentWnd,
			_Out_opt_ CCSFLAGS* lpConvertFlags,
			_Out_opt_ bool* pDoAdrBook,
			_Out_opt_ bool* pDoApply,
			_Out_opt_ HCHARSET* phCharSet,
			_Out_opt_ CSETAPPLYTYPE* pcSetApplyType,
			_Out_opt_ bool* pbUnicode)
		{
			if (lpConvertFlags) *lpConvertFlags = {};
			if (pDoAdrBook) *pDoAdrBook = {};
			if (pDoApply) *pDoApply = {};
			if (phCharSet) *phCharSet = {};
			if (pcSetApplyType) *pcSetApplyType = {};
			if (pbUnicode) *pbUnicode = {};
			if (!lpConvertFlags || !pDoAdrBook || !pDoApply || !phCharSet || !pcSetApplyType)
				return MAPI_E_INVALID_PARAMETER;
			auto hRes = S_OK;

			dialog::editor::CEditor MyData(
				pParentWnd, IDS_CONVERTFROMEML, IDS_CONVERTFROMEMLPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_CONVERTFLAGS, false));
			MyData.SetHex(0, CCSF_SMTP);
			MyData.AddPane(viewpane::CheckPane::Create(1, IDS_CONVERTCODEPAGE, false, false));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_CONVERTCODEPAGE, false));
			MyData.SetDecimal(2, CP_USASCII);
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(3, IDS_CONVERTCHARSETTYPE, false));
			MyData.SetDecimal(3, CHARSET_BODY);
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_CONVERTCHARSETAPPLYTYPE, false));
			MyData.SetDecimal(4, CSET_APPLY_UNTAGGED);
			MyData.AddPane(viewpane::CheckPane::Create(5, IDS_CONVERTDOADRBOOK, false, false));
			if (pbUnicode)
			{
				MyData.AddPane(viewpane::CheckPane::Create(6, IDS_SAVEUNICODE, false, false));
			}

			if (!MyData.DisplayDialog()) return MAPI_E_USER_CANCEL;

			*lpConvertFlags = static_cast<CCSFLAGS>(MyData.GetHex(0));
			if (MyData.GetCheck(1))
			{
				if (SUCCEEDED(hRes)) *pDoApply = true;
				*pcSetApplyType = static_cast<CSETAPPLYTYPE>(MyData.GetDecimal(4));
				if (*pcSetApplyType > CSET_APPLY_TAG_ALL) return MAPI_E_INVALID_PARAMETER;
				const auto ulCodePage = MyData.GetDecimal(2);
				const auto cCharSetType = static_cast<CHARSETTYPE>(MyData.GetDecimal(3));
				if (cCharSetType > CHARSET_WEB) return MAPI_E_INVALID_PARAMETER;
				hRes = EC_H(import::MyMimeOleGetCodePageCharset(ulCodePage, cCharSetType, phCharSet));
			}

			*pDoAdrBook = MyData.GetCheck(5);
			if (pbUnicode)
			{
				*pbUnicode = MyData.GetCheck(6);
			}

			return hRes;
		}

		// Adds menu items appropriate to the context
		// Returns number of menu items added
		_Check_return_ ULONG ExtendAddInMenu(HMENU hMenu, ULONG ulAddInContext)
		{
			output::DebugPrint(output::DBGAddInPlumbing, L"Extending menus, ulAddInContext = 0x%08X\n", ulAddInContext);
			HMENU hAddInMenu = nullptr;

			UINT uidCurMenu = ID_ADDINMENU;

			if (MENU_CONTEXT_PROPERTY == ulAddInContext)
			{
				uidCurMenu = ID_ADDINPROPERTYMENU;
			}

			for (const auto& addIn : g_lpMyAddins)
			{
				output::DebugPrint(output::DBGAddInPlumbing, L"Examining add-in for menus\n");
				if (addIn.hMod)
				{
					auto hRes = S_OK;
					if (addIn.szName)
					{
						output::DebugPrint(output::DBGAddInPlumbing, L"Examining \"%ws\"\n", addIn.szName);
					}

					for (ULONG ulMenu = 0; ulMenu < addIn.ulMenu && SUCCEEDED(hRes); ulMenu++)
					{
						if (addIn.lpMenu[ulMenu].ulFlags & MENU_FLAGS_SINGLESELECT &&
							addIn.lpMenu[ulMenu].ulFlags & MENU_FLAGS_MULTISELECT)
						{
							// Invalid combo of flags - don't add the menu
							output::DebugPrint(
								output::DBGAddInPlumbing,
								L"Invalid flags on menu \"%ws\" in add-in \"%ws\"\n",
								addIn.lpMenu[ulMenu].szMenu,
								addIn.szName);
							output::DebugPrint(
								output::DBGAddInPlumbing,
								L"MENU_FLAGS_SINGLESELECT and MENU_FLAGS_MULTISELECT cannot be combined\n");
							continue;
						}

						if (addIn.lpMenu[ulMenu].ulContext & ulAddInContext)
						{
							// Add the Add-Ins menu if we haven't added it already
							if (!hAddInMenu)
							{
								hAddInMenu = CreatePopupMenu();
								if (hAddInMenu)
								{
									InsertMenuW(
										hMenu,
										static_cast<UINT>(-1),
										MF_BYPOSITION | MF_POPUP | MF_ENABLED,
										reinterpret_cast<UINT_PTR>(hAddInMenu),
										strings::loadstring(IDS_ADDINSMENU).c_str());
								}
								else
									continue;
							}

							// Now add each of the menu entries
							if (SUCCEEDED(hRes))
							{
								const auto lpMenu = CreateMenuEntry(addIn.lpMenu[ulMenu].szMenu);
								if (lpMenu)
								{
									lpMenu->m_AddInData = reinterpret_cast<ULONG_PTR>(&addIn.lpMenu[ulMenu]);
								}

								hRes = EC_B(AppendMenuW(
									hAddInMenu,
									MF_ENABLED | MF_OWNERDRAW,
									uidCurMenu,
									reinterpret_cast<LPCWSTR>(lpMenu)));
								uidCurMenu++;
							}
						}
					}
				}
			}

			output::DebugPrint(output::DBGAddInPlumbing, L"Done extending menus\n");
			return uidCurMenu - ID_ADDINMENU;
		}

		_Check_return_ LPMENUITEM GetAddinMenuItem(HWND hWnd, UINT uidMsg)
		{
			if (uidMsg < ID_ADDINMENU) return nullptr;

			MENUITEMINFOW subMenu = {};
			subMenu.cbSize = sizeof(MENUITEMINFO);
			subMenu.fMask = MIIM_STATE | MIIM_ID | MIIM_DATA;

			if (GetMenuItemInfoW(GetMenu(hWnd), uidMsg, false, &subMenu) && subMenu.dwItemData)
			{
				return reinterpret_cast<LPMENUITEM>(reinterpret_cast<LPMENUENTRY>(subMenu.dwItemData)->m_AddInData);
			}

			return nullptr;
		}

		void displayError(const std::wstring& errString)
		{
			dialog::editor::CEditor Err(nullptr, ID_PRODUCTNAME, NULL, CEDITOR_BUTTON_OK);
			Err.SetPromptPostFix(errString);
			(void) Err.DisplayDialog();
		}

		void OutputToDbgView(const std::wstring& szMsg) { dialog::editor::OutputToDbgView(szMsg); }

		_Check_return_ LPMDB OpenMailboxWithPrompt(
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMDB lpMDB,
			const std::string& szServerName,
			const std::wstring& szMailboxDN,
			ULONG ulFlags) // desired flags for CreateStoreEntryID
		{
			if (!lpMAPISession) return nullptr;

			dialog::editor::CEditor MyPrompt(
				nullptr, IDS_OPENOTHERUSER, IDS_OPENWITHFLAGSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyPrompt.SetPromptPostFix(flags::AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
			MyPrompt.AddPane(viewpane::TextPane::CreateSingleLinePane(
				0, IDS_SERVERNAME, strings::stringTowstring(szServerName), false));
			MyPrompt.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_USERDN, szMailboxDN, false));
			MyPrompt.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_USER_SMTP_ADDRESS, false));
			MyPrompt.AddPane(viewpane::TextPane::CreateSingleLinePane(3, IDS_CREATESTORENTRYIDFLAGS, false));
			MyPrompt.SetHex(3, ulFlags);
			MyPrompt.AddPane(viewpane::CheckPane::Create(4, IDS_FORCESERVER, false, false));
			if (!MyPrompt.DisplayDialog()) return nullptr;

			return mapi::store::OpenOtherUsersMailbox(
				lpMAPISession,
				lpMDB,
				strings::wstringTostring(MyPrompt.GetStringW(0)),
				strings::wstringTostring(MyPrompt.GetStringW(1)),
				MyPrompt.GetStringW(2),
				MyPrompt.GetHex(3),
				MyPrompt.GetCheck(4));
		}

		// Display a UI to select a mailbox, then call OpenOtherUsersMailbox with the mailboxDN
		// May return MAPI_E_CANCEL
		_Check_return_ LPMDB OpenOtherUsersMailboxFromGal(_In_ LPMAPISESSION lpMAPISession, _In_ LPADRBOOK lpAddrBook)
		{
			if (!lpMAPISession || !lpAddrBook) return nullptr;

			LPMDB lpOtherUserMDB = nullptr;

			const auto szServerName = mapi::store::GetServerName(lpMAPISession);

			auto lpPrivateMDB = mapi::store::OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid);
			if (lpPrivateMDB && mapi::store::StoreSupportsManageStore(lpPrivateMDB))
			{
				auto lpMailUser = mapi::ab::SelectUser(lpAddrBook, GetDesktopWindow(), nullptr);
				if (lpMailUser)
				{
					LPSPropValue lpEmailAddress = nullptr;
					EC_MAPI_S(HrGetOneProp(lpMailUser, PR_EMAIL_ADDRESS_W, &lpEmailAddress));
					if (strings::CheckStringProp(lpEmailAddress, PT_UNICODE))
					{
						lpOtherUserMDB = OpenMailboxWithPrompt(
							lpMAPISession,
							lpPrivateMDB,
							szServerName,
							lpEmailAddress->Value.lpszW,
							OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP);
					}

					MAPIFreeBuffer(lpEmailAddress);
					lpMailUser->Release();
				}
			}

			if (lpPrivateMDB) lpPrivateMDB->Release();
			return lpOtherUserMDB;
		}

		std::string PromptServerName()
		{
			// prompt the user to enter a server name
			dialog::editor::CEditor MyData(
				nullptr, IDS_SERVERNAME, IDS_SERVERNAMEMISSINGPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SERVERNAME, false));

			if (MyData.DisplayDialog())
			{
				return strings::wstringTostring(MyData.GetStringW(0));
			}

			return std::string{};
		};

	} // namespace mapiui
} // namespace ui