// Stand alone MAPI profile functions
#pragma once

namespace ui
{
	namespace profile
	{
		std::wstring
		LaunchProfileWizard(_In_ HWND hParentWnd, ULONG ulFlags, _In_ const std::string& szServiceNameToAdd);

		void DisplayMAPISVCPath(_In_ CWnd* pParentWnd);
		std::wstring GetMAPISVCPath();

		void AddServicesToMapiSvcInf();
		void RemoveServicesFromMapiSvcInf();
	} // namespace profile
} // namespace ui