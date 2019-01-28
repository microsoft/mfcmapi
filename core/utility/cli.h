#pragma once

// MrMAPI command line
namespace cli
{
#define ulNoMatch 0xffffffff

	template <typename T> struct COMMANDLINE_SWITCH
	{
		T iSwitch;
		LPCWSTR szSwitch;
	};

	template <typename S, typename M, typename O> struct OptParser
	{
		S Switch{};
		M Mode{};
		int MinArgs{};
		int MaxArgs{};
		O ulOpt{};
	};

	template <typename S, typename M, typename O>
	OptParser<S, M, O> GetParser(S Switch, const std::vector<OptParser<S, M, O>>& parsers)
	{
		for (auto& parser : parsers)
		{
			if (Switch == parser.Switch) return parser;
		}

		return {};
	}

	std::vector<std::wstring> GetCommandLine(_In_ int argc, _In_count_(argc) const char* const argv[]);
} // namespace cli