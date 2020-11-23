#pragma once

namespace file
{
	constexpr int MAXSUBJTIGHT = 10;
	constexpr int MAXBIN = 141;
	constexpr int MAXEXT = 4;
	constexpr int MAXATTACH = 10;
	constexpr int MAXMSGPATH = MAX_PATH - MAXSUBJTIGHT - MAXBIN - MAXEXT;

	std::wstring GetSystemDirectory();
	std::wstring ShortenPath(const std::wstring& path);
	std::wstring GetDirectoryPath(HWND hWnd);

	std::wstring BuildFileNameAndPath(
		_In_ const std::wstring& szExt,
		_In_ const std::wstring& szSubj,
		_In_ const std::wstring& szRootPath,
		_In_opt_ const _SBinary* lpBin);

	std::wstring GetModuleFileName(_In_opt_ HMODULE hModule);
	std::map<std::wstring, std::wstring> GetFileVersionInfo(_In_opt_ HMODULE hModule);
} // namespace file