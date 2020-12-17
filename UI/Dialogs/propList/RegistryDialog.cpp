#include <StdAfx.h>
#include <UI/Dialogs/propList/RegistryDialog.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
//#include <UI/Dialogs/MFCUtilityFunctions.h>
//#include <core/utility/strings.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/output.h>
#include <core/propertyBag/registryPropertyBag.h>

namespace dialog
{
	static std::wstring CLASS = L"RegistryDialog";

	RegistryDialog::RegistryDialog(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects)
		: CBaseDialog(pParentWnd, lpMapiObjects, NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_szTitle = strings::loadstring(IDS_REGISTRY_DIALOG);

		CBaseDialog::CreateDialogAndMenu(NULL, NULL, NULL);
	}

	RegistryDialog::~RegistryDialog() { TRACE_DESTRUCTOR(CLASS); }

	BOOL RegistryDialog::OnInitDialog()
	{
		const auto bRet = CBaseDialog::OnInitDialog();

		UpdateTitleBarText();

		if (m_lpFakeSplitter)
		{
			m_lpRegKeyTree.Create(&m_lpFakeSplitter, true);
			m_lpFakeSplitter.SetPaneOne(m_lpRegKeyTree.GetSafeHwnd());
			m_lpRegKeyTree.FreeNodeDataCallback = [&](auto hKey) {
				if (hKey) WC_W32_S(RegCloseKey(reinterpret_cast<HKEY>(hKey)));
			};
			m_lpRegKeyTree.ItemSelectedCallback = [&](auto hItem) {
				auto hKey = reinterpret_cast<HKEY>(m_lpRegKeyTree.GetItemData(hItem));
				m_lpPropDisplay->SetDataSource(std::make_shared<propertybag::registryPropertyBag>(hKey));
			};

			EnumRegistry();

			m_lpFakeSplitter.SetPercent(0.25);
		}

		return bRet;
	}

	static const wchar_t* profileWMS =
		L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Windows Messaging Subsystem\\Profiles";
	static const wchar_t* profile15 = L"SOFTWARE\\Microsoft\\Office\\15.0\\Outlook\\Profiles";
	static const wchar_t* profile16 = L"SOFTWARE\\Microsoft\\Office\\16.0\\Outlook\\Profiles";
	void RegistryDialog::EnumRegistry()
	{
		AddProfileRoot(L"WMS", profileWMS);
		AddProfileRoot(L"15", profile15);
		AddProfileRoot(L"16", profile16);
	}

	void RegistryDialog::AddProfileRoot(const std::wstring& szName, const std::wstring& szRoot)
	{
		HKEY hRootKey = nullptr;
		auto hRes = WC_W32(RegOpenKeyExW(HKEY_CURRENT_USER, szRoot.c_str(), NULL, KEY_READ, &hRootKey));

		if (SUCCEEDED(hRes))
		{
			const auto rootNode = m_lpRegKeyTree.AddChildNode(szName.c_str(), TVI_ROOT, hRootKey, nullptr);
			AddChildren(rootNode, hRootKey);
		}
	}

	void RegistryDialog::AddChildren(const HTREEITEM hParent, const HKEY hRootKey)
	{
		auto cchMaxSubKeyLen = DWORD{}; // Param in RegQueryInfoKeyW is misnamed
		auto cSubKeys = DWORD{};
		auto hRes = WC_W32(RegQueryInfoKeyW(
			hRootKey,
			nullptr, // lpClass
			nullptr, // lpcchClass
			nullptr, // lpReserved
			&cSubKeys, // lpcSubKeys
			&cchMaxSubKeyLen, // lpcbMaxSubKeyLen
			nullptr, // lpcbMaxClassLen
			nullptr, // lpcValues
			nullptr, // lpcbMaxValueNameLen
			nullptr, // lpcbMaxValueLen
			nullptr, // lpcbSecurityDescriptor
			nullptr)); // lpftLastWriteTime

		if (cSubKeys && cchMaxSubKeyLen)
		{
			cchMaxSubKeyLen++; // For null terminator
			auto szBuf = std::wstring(cchMaxSubKeyLen, '\0');
			for (DWORD dwIndex = 0; dwIndex < cSubKeys; dwIndex++)
			{
				szBuf.clear();
				hRes = WC_W32(RegEnumKeyW(hRootKey, dwIndex, const_cast<wchar_t*>(szBuf.c_str()), cchMaxSubKeyLen));
				if (hRes == S_OK)
				{
					HKEY hSubKey = nullptr;
					hRes = WC_W32(RegOpenKeyExW(hRootKey, szBuf.c_str(), NULL, KEY_READ, &hSubKey));
					const auto node = m_lpRegKeyTree.AddChildNode(szBuf.c_str(), hParent, hSubKey, nullptr);
					AddChildren(node, hSubKey);
				}
			}
		}
	}

	BEGIN_MESSAGE_MAP(RegistryDialog, CBaseDialog)
	END_MESSAGE_MAP()
} // namespace dialog
