#pragma once
#include <UI/Dialogs/BaseDialog.h>
#include <UI/Controls/StyleTree/StyleTreeCtrl.h>
#include <core/sortlistdata/sortListData.h>

namespace dialog
{
	class RegistryDialog : public CBaseDialog
	{
	public:
		RegistryDialog(_In_ ui::CParentWnd* pParentWnd, _In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects);
		~RegistryDialog();

	protected:
		// Overrides from base class
		BOOL OnInitDialog() override;

	private:
		controls::StyleTreeCtrl m_lpRegKeyList{};

		void EnumAccounts(const std::wstring& szCat, const CLSID* pclsidCategory);
		void EnumRegistry();
		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog