#include <core/stdafx.h>
#include <core/utility/cli.h>
#include <core/utility/strings.h>

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

	// Parses command line arguments and fills out OPTIONS
	void ParseArgs(
		OPTIONS& options,
		std::deque<std::wstring>& args,
		const std::vector<OptParser>& _parsers,
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
			if (!opt.CheckMinArgs(args, _parsers))
			{
				// resetting our mode here, switch to help
				options.mode = cmdmodeHelp;
			}

			if (!DoSwitch(&options, opt, args))
			{
				options.mode = cmdmodeHelp;
			}
		}

		postParseCheck(&options);
	}

	// Return true if we succesfully peeled off a switch.
	// Return false on an error.
	_Check_return_ bool DoSwitch(OPTIONS* options, const cli::OptParser& opt, std::deque<std::wstring>& args)
	{
		const auto arg0 = args.front();
		args.pop_front();

		if (opt.parseArgs) return opt.parseArgs(options, args);

		switch (opt.clSwitch)
		{
		case switchNoSwitch:
			// naked option without a flag - we only allow one of these
			if (!options->lpszUnswitchedOption.empty())
			{
				return false;
			} // He's already got one, you see.

			options->lpszUnswitchedOption = arg0;
			break;
		case switchUnknown:
			// display help
			return false;
		}

		return true;
	}

	_Check_return_ bool
	OptParser::CheckMinArgs(const std::deque<std::wstring>& args, const std::vector<OptParser>& _parsers) const
	{
		if (minArgs == 0) return true;
		if (args.size() <= minArgs) return false;

		auto c = UINT{0};
		for (auto it = args.cbegin() + 1; it != args.cend() && c < minArgs; it++, c++)
		{
			if (ParseArgument(*it, _parsers) != switchNoSwitch)
			{
				return false;
			}
		}

		return true;
	}
} // namespace cli