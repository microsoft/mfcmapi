#pragma once

namespace version
{
	std::wstring GetOutlookVersionString();
	bool Is64BitModule(const std::wstring& modulePath);
} // namespace version