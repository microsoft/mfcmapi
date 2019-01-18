#include <StdAfx.h>
#include <UI/addinui.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <UI/UIFunctions.h>
#include <Addins.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace ui
{
	namespace addinui
	{
		_Check_return_ __declspec(dllexport) HRESULT
			__cdecl SimpleDialog(_In_z_ LPWSTR szTitle, _Printf_format_string_ LPWSTR szMsg, ...)
		{
			dialog::editor::CEditor MySimpleDialog(nullptr, NULL, NULL, CEDITOR_BUTTON_OK);
			MySimpleDialog.SetAddInTitle(szTitle);

			va_list argList = nullptr;
			va_start(argList, szMsg);
			const auto szDialogString = strings::formatV(szMsg, argList);
			va_end(argList);

			MySimpleDialog.SetPromptPostFix(szDialogString);

			return MySimpleDialog.DisplayDialog() ? S_OK : S_FALSE;
		}

		_Check_return_ __declspec(dllexport) HRESULT
			__cdecl ComplexDialog(_In_ LPADDINDIALOG lpDialog, _Out_ LPADDINDIALOGRESULT* lppDialogResult)
		{
			if (!lpDialog) return MAPI_E_INVALID_PARAMETER;
			// Reject any flags except CEDITOR_BUTTON_OK and CEDITOR_BUTTON_CANCEL
			if (lpDialog->ulButtonFlags & ~(CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)) return MAPI_E_INVALID_PARAMETER;
			auto hRes = S_OK;

			dialog::editor::CEditor MyComplexDialog(nullptr, NULL, NULL, lpDialog->ulButtonFlags);
			MyComplexDialog.SetAddInTitle(lpDialog->szTitle);
			MyComplexDialog.SetPromptPostFix(lpDialog->szPrompt);

			if (lpDialog->ulNumControls && lpDialog->lpDialogControls)
			{
				for (ULONG i = 0; i < lpDialog->ulNumControls; i++)
				{
					switch (lpDialog->lpDialogControls[i].cType)
					{
					case ADDIN_CTRL_CHECK:
						MyComplexDialog.AddPane(viewpane::CheckPane::Create(
							i,
							NULL,
							lpDialog->lpDialogControls[i].bDefaultCheckState,
							lpDialog->lpDialogControls[i].bReadOnly));
						break;
					case ADDIN_CTRL_EDIT_TEXT:
					{
						if (lpDialog->lpDialogControls[i].bMultiLine)
						{
							MyComplexDialog.AddPane(viewpane::TextPane::CreateCollapsibleTextPane(
								i, NULL, lpDialog->lpDialogControls[i].bReadOnly));
							MyComplexDialog.SetStringW(i, lpDialog->lpDialogControls[i].szDefaultText);
						}
						else
						{
							MyComplexDialog.AddPane(viewpane::TextPane::CreateSingleLinePane(
								i,
								NULL,
								std::wstring(lpDialog->lpDialogControls[i].szDefaultText),
								lpDialog->lpDialogControls[i].bReadOnly));
						}

						break;
					}
					case ADDIN_CTRL_EDIT_BINARY:
						if (lpDialog->lpDialogControls[i].bMultiLine)
						{
							MyComplexDialog.AddPane(viewpane::TextPane::CreateCollapsibleTextPane(
								i, NULL, lpDialog->lpDialogControls[i].bReadOnly));
						}
						else
						{
							MyComplexDialog.AddPane(viewpane::TextPane::CreateSingleLinePane(
								i, NULL, lpDialog->lpDialogControls[i].bReadOnly));
						}
						MyComplexDialog.SetBinary(
							i, lpDialog->lpDialogControls[i].lpBin, lpDialog->lpDialogControls[i].cbBin);
						break;
					case ADDIN_CTRL_EDIT_NUM_DECIMAL:
						MyComplexDialog.AddPane(
							viewpane::TextPane::CreateSingleLinePane(i, NULL, lpDialog->lpDialogControls[i].bReadOnly));
						MyComplexDialog.SetDecimal(i, lpDialog->lpDialogControls[i].ulDefaultNum);
						break;
					case ADDIN_CTRL_EDIT_NUM_HEX:
						MyComplexDialog.AddPane(
							viewpane::TextPane::CreateSingleLinePane(i, NULL, lpDialog->lpDialogControls[i].bReadOnly));
						MyComplexDialog.SetHex(i, lpDialog->lpDialogControls[i].ulDefaultNum);
						break;
					}

					// Do this after initializing controls so we have our label status set correctly.
					MyComplexDialog.SetAddInLabel(i, lpDialog->lpDialogControls[i].szLabel);
				}
			}

			if (MyComplexDialog.DisplayDialog())
			{
				// Put together results if needed
				if (lppDialogResult && lpDialog->ulNumControls && lpDialog->lpDialogControls)
				{
					const auto lpResults = new _AddInDialogResult;
					if (lpResults)
					{
						lpResults->ulNumControls = lpDialog->ulNumControls;
						lpResults->lpDialogControlResults = new _AddInDialogControlResult[lpDialog->ulNumControls];
						if (lpResults->lpDialogControlResults)
						{
							ZeroMemory(
								lpResults->lpDialogControlResults,
								sizeof(_AddInDialogControlResult) * lpDialog->ulNumControls);
							for (ULONG i = 0; i < lpDialog->ulNumControls; i++)
							{
								lpResults->lpDialogControlResults[i].cType = lpDialog->lpDialogControls[i].cType;
								switch (lpDialog->lpDialogControls[i].cType)
								{
								case ADDIN_CTRL_CHECK:
									lpResults->lpDialogControlResults[i].bCheckState = MyComplexDialog.GetCheck(i);
									break;
								case ADDIN_CTRL_EDIT_TEXT:
								{
									auto szText = MyComplexDialog.GetStringW(i);
									if (!szText.empty())
									{
										lpResults->lpDialogControlResults[i].szText =
											LPWSTR(strings::wstringToLPCWSTR(szText));
									}
									break;
								}
								case ADDIN_CTRL_EDIT_BINARY:
									// GetEntryID does just what we want - abuse it
									hRes = WC_H(MyComplexDialog.GetEntryID(
										i,
										false,
										&lpResults->lpDialogControlResults[i].cbBin,
										reinterpret_cast<LPENTRYID*>(&lpResults->lpDialogControlResults[i].lpBin)));
									break;
								case ADDIN_CTRL_EDIT_NUM_DECIMAL:
									lpResults->lpDialogControlResults[i].ulVal = MyComplexDialog.GetDecimal(i);
									break;
								case ADDIN_CTRL_EDIT_NUM_HEX:
									lpResults->lpDialogControlResults[i].ulVal = MyComplexDialog.GetHex(i);
									break;
								}
							}
						}

						if (SUCCEEDED(hRes))
						{
							*lppDialogResult = lpResults;
						}
						else
							FreeDialogResult(lpResults);
					}
				}
			}

			return hRes;
		}

		__declspec(dllexport) void __cdecl FreeDialogResult(_In_ LPADDINDIALOGRESULT lpDialogResult)
		{
			if (lpDialogResult)
			{
				if (lpDialogResult->lpDialogControlResults)
				{
					for (ULONG i = 0; i < lpDialogResult->ulNumControls; i++)
					{
						delete[] lpDialogResult->lpDialogControlResults[i].lpBin;
						delete[] lpDialogResult->lpDialogControlResults[i].szText;
					}
				}

				delete[] lpDialogResult;
			}
		}

		// Adds menu items appropriate to the context
		// Returns number of menu items added
		_Check_return_ ULONG ExtendAddInMenu(HMENU hMenu, ULONG ulAddInContext)
		{
			output::DebugPrint(DBGAddInPlumbing, L"Extending menus, ulAddInContext = 0x%08X\n", ulAddInContext);
			HMENU hAddInMenu = nullptr;

			UINT uidCurMenu = ID_ADDINMENU;

			if (MENU_CONTEXT_PROPERTY == ulAddInContext)
			{
				uidCurMenu = ID_ADDINPROPERTYMENU;
			}

			for (const auto& addIn : g_lpMyAddins)
			{
				output::DebugPrint(DBGAddInPlumbing, L"Examining add-in for menus\n");
				if (addIn.hMod)
				{
					auto hRes = S_OK;
					if (addIn.szName)
					{
						output::DebugPrint(DBGAddInPlumbing, L"Examining \"%ws\"\n", addIn.szName);
					}

					for (ULONG ulMenu = 0; ulMenu < addIn.ulMenu && SUCCEEDED(hRes); ulMenu++)
					{
						if (addIn.lpMenu[ulMenu].ulFlags & MENU_FLAGS_SINGLESELECT &&
							addIn.lpMenu[ulMenu].ulFlags & MENU_FLAGS_MULTISELECT)
						{
							// Invalid combo of flags - don't add the menu
							output::DebugPrint(
								DBGAddInPlumbing,
								L"Invalid flags on menu \"%ws\" in add-in \"%ws\"\n",
								addIn.lpMenu[ulMenu].szMenu,
								addIn.szName);
							output::DebugPrint(
								DBGAddInPlumbing,
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

			output::DebugPrint(DBGAddInPlumbing, L"Done extending menus\n");
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

		template <typename T> T GetFunction(HMODULE hMod, LPCSTR szFuncName)
		{
			return WC_D(T, reinterpret_cast<T>(GetProcAddress(hMod, szFuncName)));
		}

		void InvokeAddInMenu(_In_opt_ LPADDINMENUPARAMS lpParams)
		{
			if (!lpParams) return;
			if (!lpParams->lpAddInMenu) return;
			if (!lpParams->lpAddInMenu->lpAddIn) return;

			if (!lpParams->lpAddInMenu->lpAddIn->pfnCallMenu)
			{
				if (!lpParams->lpAddInMenu->lpAddIn->hMod) return;
				lpParams->lpAddInMenu->lpAddIn->pfnCallMenu =
					GetFunction<LPCALLMENU>(lpParams->lpAddInMenu->lpAddIn->hMod, szCallMenu);
			}

			if (!lpParams->lpAddInMenu->lpAddIn->pfnCallMenu)
			{
				output::DebugPrint(DBGAddInPlumbing, L"InvokeAddInMenu: CallMenu not found\n");
				return;
			}

			WC_H_S(lpParams->lpAddInMenu->lpAddIn->pfnCallMenu(lpParams));
		}
	} // namespace addinui
} // namespace ui