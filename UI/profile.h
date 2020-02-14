// Stand alone MAPI profile functions
#pragma once

namespace ui::profile
{
	std::wstring LaunchProfileWizard(_In_ HWND hParentWnd, ULONG ulFlags, _In_ const std::wstring& szServiceNameToAdd);

	void DisplayMAPISVCPath(_In_ CWnd* pParentWnd);
	std::wstring GetMAPISVCPath();

	void AddServicesToMapiSvcInf();
	void RemoveServicesFromMapiSvcInf();
} // namespace ui::profile