#include <core/stdafx.h>
#include <core/utility/cli.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace cli
{
	option switchUnswitched{L"", cmdmodeUnknown, 1, 1, OPT_NOOPT};
	option switchInvalid{L"", cmdmodeHelpFull, 0, 0, OPT_NOOPT};
	option switchHelp{L"?", cmdmodeHelpFull, 0, 0, OPT_INITMFC};
	option switchVerbose{L"Verbose", cmdmodeUnknown, 0, 0, OPT_INITMFC};

	// Consume min/max args and store them in the option
	_Check_return_ bool
	option::scanArgs(std::deque<std::wstring>& _args, OPTIONS& options, const std::vector<option*>& optionsArray)
	{
		clear(); // We're not "seen" until we get past this check
		if (_args.size() < minArgs) return false;

		// Add our flags and set mode
		options.flags |= flags;
		if (cmdmodeUnknown != mode && cmdmodeHelp != options.mode)
		{
			if (!bSetMode(options.mode, mode))
			{
				return false;
			}
		}

		UINT c{};
		auto foundArgs = std::vector<std::wstring>{};
		for (auto it = _args.cbegin(); it != _args.cend() && c < maxArgs; it++, c++)
		{
			// If we *do* get a parser while looking for our minargs, then we've failed
			if (GetOption(*it, optionsArray))
			{
				// If we've already gotten our minArgs, we're done
				if (c >= minArgs) break;
				return false;
			}

			// Check that our arguments are numbers if needed
			if (flags & OPT_NEEDNUM)
			{
				ULONG num{};
				if (!strings::tryWstringToUlong(num, *it, 10, true))
				{
					// If we've already gotten our minArgs, we're done
					if (c >= minArgs) break;
					return false;
				}
			}

			foundArgs.push_back(*it);
		}

		seen = true;

		// and save them locally
		args = foundArgs;

		// remove all our found arguments
		_args.erase(_args.begin(), _args.begin() + foundArgs.size());

		return seen;
	}

	// Checks if szArg is an option, and if it is, returns which option it is
	// We return the first match, so switches should be ordered appropriately
	option* GetOption(const std::wstring& szArg, const std::vector<option*>& optionsArray)
	{
		if (szArg.empty()) return nullptr;

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
			return nullptr;
		}

		for (const auto parser : optionsArray)
		{
			// If we have a match
			if (parser && strings::beginsWith(parser->name(), szSwitch))
			{
				return parser;
			}
		}

		return &switchInvalid;
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
	void ParseArgs(OPTIONS& options, std::deque<std::wstring>& args, const std::vector<option*>& optionsArray)
	{
		switchUnswitched.clear();

		if (args.empty())
		{
			options.mode = cmdmodeHelp;
		}

		// We pop at least one arg each iteration, so this while is OK.
		while (!args.empty())
		{
			auto arg0 = args.front();
			auto opt = GetOption(arg0, optionsArray);
			args.pop_front();

			if (!opt)
			{
				// If we didn't get a parser, treat this as an unswitched option.
				// We only allow one of these
				if (!switchUnswitched.empty())
				{
					options.mode = cmdmodeHelp;
				} // He's already got one, you see.
				else
				{
					// Trick switchUnswitched into scanning this path as an argument.
					auto args0 = std::deque<std::wstring>{arg0};
					switchUnswitched.scanArgs(args0, options, optionsArray);
				}

				continue;
			}

			if (opt == &switchInvalid)
			{
				options.mode = cmdmodeHelp;
				continue;
			}

			// Make sure we have the right number of args and set flags and mode
			// Commands with variable argument counts can special case themselves
			if (!opt->scanArgs(args, options, optionsArray))
			{
				// resetting our mode here, switch to help
				options.mode = cmdmodeHelp;
			}
		}
	}

	void PrintArgs(_In_ const OPTIONS& ProgOpts, const std::vector<option*>& optionsArray)
	{
		output::DebugPrint(DBGGeneric, L"Mode = %d\n", ProgOpts.mode);
		output::DebugPrint(DBGGeneric, L"options = 0x%08X\n", ProgOpts.flags);
		if (!cli::switchUnswitched.empty())
			output::DebugPrint(DBGGeneric, L"lpszUnswitchedOption = %ws\n", cli::switchUnswitched[0].c_str());

		for (const auto& option : optionsArray)
		{
			if (option->isSet())
			{
				output::DebugPrint(DBGGeneric, L"Switch: %ws\n", option->name());
			}
			else if (!option->empty())
			{
				output::DebugPrint(DBGGeneric, L"Switch: %ws (not set but has args)\n", option->name());
			}

			for (UINT i = 0; i < option->size(); i++)
			{
				output::DebugPrint(DBGGeneric, L"  arg[%d] = %ws\n", i, option->at(i).c_str());
			}
		}
	}
} // namespace cli