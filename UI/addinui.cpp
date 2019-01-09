#include <StdAfx.h>
#include <UI/addinui.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <IO/File.h>
#include <UI/UIFunctions.h>

namespace ui
{
	namespace addinui
	{
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