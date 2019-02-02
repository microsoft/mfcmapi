#pragma once
#include <core/utility/strings.h>

// MrMAPI command line
namespace cli
{
	template <typename T> struct COMMANDLINE_SWITCH
	{
		T iSwitch;
		LPCWSTR szSwitch;
	};

	template <typename S, typename M, typename O> struct OptParser
	{
		S clSwitch{};
		M mode{};
		int minArgs{};
		int maxArgs{};
		O options{};
	};

	template <typename S, typename M, typename O>
	OptParser<S, M, O> GetParser(S clSwitch, const std::vector<OptParser<S, M, O>>& parsers)
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
	template <typename S> S ParseArgument(const std::wstring& szArg, const std::vector<COMMANDLINE_SWITCH<S>>& switches)
	{
		if (szArg.empty()) return S(0);

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
			return S(0);
		}

		for (const auto& s : switches)
		{
			// If we have a match
			if (strings::beginsWith(s.szSwitch, szSwitch))
			{
				return s.iSwitch;
			}
		}

		return S(0);
	}

	std::vector<std::wstring> GetCommandLine(_In_ int argc, _In_count_(argc) const char* const argv[]);
} // namespace cli