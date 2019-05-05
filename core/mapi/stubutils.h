#pragma once

namespace mapistub
{
	HMODULE GetMAPIHandle();
	void UnloadPrivateMAPI();
	void ForceOutlookMAPI(bool fForce);
	void ForceSystemMAPI(bool fForce);
	void SetMAPIHandle(HMODULE hinstMAPI);
	HMODULE GetPrivateMAPI();
	std::wstring GetComponentPath(const std::wstring& szComponent, const std::wstring& szQualifier, bool fInstall);
	extern WCHAR g_pszOutlookQualifiedComponents[][MAX_PATH];
	std::vector<std::wstring> GetMAPIPaths();
	// Looks up Outlook's path given its qualified component guid
	std::wstring GetOutlookPath(_In_ const std::wstring& szCategory, _Out_opt_ bool* lpb64);
	std::wstring GetInstalledOutlookMAPI(int iOutlook);
	std::wstring GetMAPISystemDir();
} // namespace mapistub
