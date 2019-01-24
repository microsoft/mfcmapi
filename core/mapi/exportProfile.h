#pragma once

namespace output
{
	void ExportProfile(
		_In_ const std::string& szProfile,
		_In_ const std::wstring& szProfileSection,
		bool bByteSwapped,
		const std::wstring& szFileName);
}