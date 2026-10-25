// Stand alone MAPI profile functions
#pragma once

namespace ui::profile
{

	void DisplayMAPISVCPath(_In_ CWnd* pParentWnd);
	std::wstring GetMAPISVCPath();

	void AddServicesToMapiSvcInf();
	void RemoveServicesFromMapiSvcInf();
} // namespace ui::profile