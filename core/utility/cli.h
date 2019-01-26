#pragma once

// MrMAPI command line
namespace cli
{
#define ulNoMatch 0xffffffff

	std::vector<std::wstring> GetCommandLine(_In_ int argc, _In_count_(argc) const char* const argv[]);
} // namespace cli