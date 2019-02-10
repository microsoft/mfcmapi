#include <core/stdafx.h>
#include <core/utility/cli.h>
#include <core/utility/strings.h>
#include <queue>

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

	// If the mode isn't set (is 0), then we can set it to any mode
	// If the mode IS set (non 0), then we can only set it to the same mode
	// IE trying to change the mode from anything but unset will fail
	bool bSetMode(_In_ int& pMode, _In_ int targetMode)
	{
		if (0 == pMode || targetMode == pMode)
		{
			pMode = targetMode;
			return true;
		}

		return false;
	}

	// Converts an argc/argv style command line to a vector
	std::deque<std::wstring> GetCommandLine(_In_ int argc, _In_count_(argc) const char* const argv[])
	{
		auto args = std::deque<std::wstring>{};

		for (auto i = 1; i < argc; i++)
		{
			args.emplace_back(strings::LPCSTRToWstring(argv[i]));
		}

		return args;
	}

	// Assumes that our no switch case is 0
	_Check_return_ bool CheckMinArgs(
		const cli::OptParser& opt,
		const std::deque<std::wstring>& args,
		const std::vector<COMMANDLINE_SWITCH>& switches)
	{
		if (opt.minArgs == 0) return true;
		if (args.size() <= opt.minArgs) return false;

		auto c = UINT{0};
		for (auto it = args.cbegin() + 1; it != args.cend() && c < opt.minArgs; it++, c++)
		{
			if (ParseArgument(*it, switches) != 0)
			{
				return false;
			}
		}

		return true;
	}
} // namespace cli