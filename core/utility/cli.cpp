#include <core/stdafx.h>
#include <core/utility/cli.h>
#include <core/utility/strings.h>

namespace cli
{
	OptParser GetParser(int clSwitch, const std::vector<OptParser>& parsers)
	{
		for (const auto& parser : parsers)
		{
			if (clSwitch == parser.clSwitch) return parser;
		}

		return {};
	}

	// Checks if szArg is an option, and if it is, returns which option it is
	// We return the first match, so switches should be ordered appropriately
	// The first switch should be our "no match" switch
	int ParseArgument(const std::wstring& szArg, const std::vector<COMMANDLINE_SWITCH>& switches)
	{
		if (szArg.empty()) return 0;

		auto szSwitch = std::wstring{};

		// Check if this is a switch at all
		switch (szArg[0])
		{
		case L'-':
		case L'/':
		case L'\\':
			if (szArg[1] != 0) szSwitch = strings::wstringToLower(&szArg[1]);
			break;
		default:
			return 0;
		}

		for (const auto& s : switches)
		{
			// If we have a match
			if (strings::beginsWith(s.szSwitch, szSwitch))
			{
				return s.iSwitch;
			}
		}

		return 0;
	}

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