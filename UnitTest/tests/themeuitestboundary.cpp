#include <StdAfx.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/BaseDialog.h>
#include <CppUnitTest.h>

// Catalog tests must not invoke the application's window-management boundary.
// Use runtime assertions, not noreturn Fail, so Release optimization can compile the callers.
namespace ui
{
	void ApplyThemeToAllWindows() noexcept
	{
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::IsTrue(
			false, L"Catalog operation attempted to repaint windows.");
	}

	uiTheme GetEffectiveTheme() noexcept
	{
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::IsTrue(
			false, L"Catalog operation queried application appearance.");
		return uiTheme::Light;
	}

	const theme::ThemeDefinition& GetActiveTheme() noexcept
	{
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::IsTrue(
			false, L"Catalog operation queried the applied theme.");
		return theme::GetDefaultTheme();
	}
} // namespace ui

namespace dialog
{
	void CBaseDialog::UpdateStatus(HWND, statusPane, const std::wstring&) noexcept
	{
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::IsTrue(
			false, L"Catalog operation attempted to update status.");
	}
} // namespace dialog
