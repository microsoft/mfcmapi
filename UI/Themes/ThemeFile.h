#pragma once
#include <UI/Themes/Theme.h>

namespace ui::theme
{
	struct ThemeFileError
	{
		std::wstring path;
		std::wstring message;
	};

	struct ThemeCatalog
	{
		std::vector<ThemeDefinition> themes;
		std::vector<ThemeFileError> errors;
	};

	bool ParseThemeFile(std::wstring_view json, ThemeDefinition& definition, std::wstring& error);
	ThemeCatalog LoadThemeDirectory(const std::wstring& directory);
	const std::vector<ThemeFileError>& GetThemeFileErrors();
} // namespace ui::theme
