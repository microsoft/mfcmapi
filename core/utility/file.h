#pragma once

namespace file
{
	std::wstring ShortenPath(const std::wstring& path);
	std::wstring GetDirectoryPath(HWND hWnd);

#define MAXSUBJ 25
#define MAXSUBJTIGHT 10
#define MAXBIN 141
#define MAXEXT 4
#define MAXATTACH 10
#define MAXMSGPATH (MAX_PATH - MAXSUBJTIGHT - MAXBIN - MAXEXT)
	std::wstring BuildFileNameAndPath(
		_In_ const std::wstring& szExt,
		_In_ const std::wstring& szSubj,
		_In_ const std::wstring& szRootPath,
		_In_opt_ const _SBinary* lpBin);

	std::wstring GetModuleFileName(_In_opt_ HMODULE hModule);
	std::map<std::wstring, std::wstring> GetFileVersionInfo(_In_opt_ HMODULE hModule);
} // namespace file