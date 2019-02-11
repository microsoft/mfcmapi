#include <core/stdafx.h>
#include <core/utility/cli.h>
#include <core/utility/strings.h>
#include <queue>

namespace cli
{
	const OptParser helpParser = {switchHelp, L"?", cmdmodeHelpFull, 0, 0, OPT_INITMFC};
	const OptParser verboseParser = {switchVerbose, L"Verbose", cmdmodeUnknown, 0, 0, OPT_VERBOSE | OPT_INITMFC};
	const OptParser noSwitchParser = {switchNoSwitch, L"", cmdmodeUnknown, 0, 0, OPT_NOOPT};
	const std::vector<OptParser> parsers = {helpParser, verboseParser, noSwitchParser};

	OptParser GetParser(int clSwitch, const std::vector<OptParser>& _parsers)
	{
		for (const auto& parser : _parsers)
		{
			if (clSwitch == parser.clSwitch) return parser;
		}

		return {};
	}

	// Checks if szArg is an option, and if it is, returns which option it is
	// We return the first match, so switches should be ordered appropriately
	int ParseArgument(const std::wstring& szArg, const std::vector<OptParser>& _parsers)
	{
		if (szArg.empty()) return switchEnum::switchNoSwitch;

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
			return switchEnum::switchNoSwitch;
		}

		for (const auto& s : _parsers)
		{
			// If we have a match
			if (strings::beginsWith(s.szSwitch, szSwitch))
			{
				return s.clSwitch;
			}
		}

		return switchEnum::switchNoSwitch;
	}

	// If the mode isn't set (is cmdmodeUnknown/0), then we can set it to any mode
	// If the mode IS set (non cmdmodeUnknown/0), then we can only set it to the same mode
	// IE trying to change the mode from anything but unset will fail
	bool bSetMode(_In_ int& pMode, _In_ int targetMode)
	{
		if (cmdmodeUnknown == pMode || targetMode == pMode)
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

	_Check_return_ bool CheckMinArgs(
		const cli::OptParser& opt,
		const std::deque<std::wstring>& args,
		const std::vector<OptParser>& _parsers)
	{
		if (opt.minArgs == 0) return true;
		if (args.size() <= opt.minArgs) return false;

		auto c = UINT{0};
		for (auto it = args.cbegin() + 1; it != args.cend() && c < opt.minArgs; it++, c++)
		{
			if (ParseArgument(*it, _parsers) != switchNoSwitch)
			{
				return false;
			}
		}

		return true;
	}

	// Parses command line arguments and fills out OPTIONS
	void ParseArgs(
		OPTIONS& options,
		std::deque<std::wstring>& args,
		const std::vector<OptParser>& _parsers,
		std::function<bool(OPTIONS* _options, const cli::OptParser& opt, std::deque<std::wstring>& args)> doSwitch,
		std::function<void(OPTIONS* _options)> postParseCheck)
	{
		if (args.empty())
		{
			options.mode = cmdmodeHelp;
		}

		// DoSwitch will either consume part of args or return an error, so this while is OK.
		while (!args.empty())
		{
			const auto iSwitch = ParseArgument(args.front(), _parsers);
			const auto opt = GetParser(iSwitch, _parsers);
			if (opt.mode == cmdmodeHelpFull)
			{
				options.mode = cmdmodeHelpFull;
			}

			options.options |= opt.options;
			if (cmdmodeUnknown != opt.mode && cmdmodeHelp != options.mode)
			{
				if (!bSetMode(options.mode, opt.mode))
				{
					// resetting our mode here, switch to help
					options.mode = cmdmodeHelp;
				}
			}

			// Make sure we have the minimum number of args
			// Commands with variable argument counts can special case themselves
			if (!CheckMinArgs(opt, args, _parsers))
			{
				// resetting our mode here, switch to help
				options.mode = cmdmodeHelp;
			}

			if (!doSwitch(&options, opt, args))
			{
				options.mode = cmdmodeHelp;
			}
		}

		postParseCheck(&options);
	}
} // namespace cli