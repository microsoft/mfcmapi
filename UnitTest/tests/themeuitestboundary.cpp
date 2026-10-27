#include <StdAfx.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/BaseDialog.h>
#include <CppUnitTest.h>

// Catalog tests must not invoke the application's window-management boundary.
namespace ui
{
	void ApplyThemeToAllWindows() noexcept
	{
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::Fail(L"Catalog operation attempted to repaint windows.");
	}

	uiTheme GetEffectiveTheme() noexcept
	{
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::Fail(
			L"Catalog operation queried application appearance.");
	}
} // namespace ui

namespace dialog
{
	void CBaseDialog::UpdateStatus(HWND, statusPane, const std::wstring&) noexcept
	{
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::Fail(L"Catalog operation attempted to update status.");
	}
} // namespace dialog
