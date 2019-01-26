#include <core/stdafx.h>
#include <core/utility/cli.h>
#include <core/utility/strings.h>

namespace cli
{
	// Converts an argc/argv style command line to a vector
	std::vector<std::wstring> GetCommandLine(_In_ int argc, _In_count_(argc) const char* const argv[])
	{
		auto args = std::vector<std::wstring>{};

		for (auto i = 1; i < argc; i++)
		{
			args.emplace_back(strings::LPCSTRToWstring(argv[i]));
		}

		return args;
	}
} // namespace cli