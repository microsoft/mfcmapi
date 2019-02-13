#pragma once

// MrMAPI command line
namespace cli
{
	struct OPTIONS
	{
		int mode{0};
		int options{0};
		std::wstring lpszUnswitchedOption;
		// dummy function to force a vtable so we can use dynamic_cast
		virtual void __dummy() noexcept {};
	};

	class OptParser
	{
	private:
		bool seen{false};

	public:
		int clSwitch{};
		LPCWSTR szSwitch{};
		int mode{};
		UINT minArgs{};
		UINT maxArgs{};
		int options{};
		std::function<bool(OPTIONS* _options, std::deque<std::wstring>& args)> parseArgs = 0;

		OptParser() = default;

		OptParser(
			int _clSwitch,
			LPCWSTR _szSwitch,
			int _mode,
			UINT _minArgs,
			UINT _maxArgs,
			int _options,
			std::function<bool(OPTIONS* _options, std::deque<std::wstring>& args)> _parseArgs = 0)
		{
			clSwitch = _clSwitch;
			szSwitch = _szSwitch;
			mode = _mode;
			minArgs = _minArgs;
			maxArgs = _maxArgs;
			options = _options;
			parseArgs = _parseArgs;
		}

		void SetSeen(bool _seen) { seen = _seen; }
		operator bool() const { return seen; }

		_Check_return_ bool
		CheckMinArgs(const std::deque<std::wstring>& args, const std::vector<OptParser>& _parsers) const;
	};

	enum switchEnum
	{
		switchNoSwitch = 0, // not a switch
		switchUnknown, // unknown switch
		switchHelp, // '-h'
		switchVerbose, // '-v'
		switchFirstSwitch, // When extending switches, use this as the value of the first switch
	};

	enum modeEnum
	{
		cmdmodeUnknown = 0,
		cmdmodeHelp,
		cmdmodeHelpFull,
		cmdmodeFirstMode, // When extending modes, use this as the value of the first mode
	};

	enum flagsEnum
	{
		OPT_NOOPT = 0x00000,
		OPT_VERBOSE = 0x00001,
		OPT_INITMFC = 0x00002,
	};

	extern const OptParser helpParser;
	extern const OptParser verboseParser;
	extern const OptParser noSwitchParser;

	extern const std::vector<OptParser> parsers;

	OptParser GetParser(int clSwitch, const std::vector<OptParser>& parsers);

	// Checks if szArg is an option, and if it is, returns which option it is
	// We return the first match, so switches should be ordered appropriately
	// The first switch should be our "no match" switch
	int ParseArgument(const std::wstring& szArg, const std::vector<OptParser>& _parsers);

	// If the mode isn't set (is 0), then we can set it to any mode
	// If the mode IS set (non 0), then we can only set it to the same mode
	// IE trying to change the mode from anything but unset will fail
	bool bSetMode(_In_ int& pMode, _In_ int targetMode);

	std::deque<std::wstring> GetCommandLine(_In_ int argc, _In_count_(argc) const char* const argv[]);

	// Parses command line arguments and fills out OPTIONS
	void ParseArgs(
		OPTIONS& options,
		std::deque<std::wstring>& args,
		const std::vector<OptParser>& parsers,
		std::function<void(OPTIONS* _options)> postParseCheck);

	_Check_return_ bool DoSwitch(OPTIONS* _options, const cli::OptParser& opt, std::deque<std::wstring>& args);
} // namespace cli