#pragma once
#include <UI/Dialogs/Editors/Editor.h>
#include <array>

namespace dialog::editor
{
	enum class MailboxMethod
	{
		SMTP,
		DN,
		EntryID
	};

	struct MailboxOptions
	{
		std::wstring identity;
		std::wstring server;
		std::wstring smtp;
		ULONG flags{};
		bool forceServer{};
		bool base64{};
		bool displayProps{};
		bool unwrap{};
	};

	class CMailboxEditor : public CEditor
	{
	public:
		CMailboxEditor(_In_ CWnd* parent, _In_ LPMAPISESSION session);
		MailboxMethod GetMethod() const noexcept { return m_method; }
		const MailboxOptions& GetOptions() const noexcept { return m_options[static_cast<size_t>(m_method)]; }

	private:
		BOOL OnInitDialog() override;
		ULONG HandleChange(UINT nID) override;
		void OnEditAction1() override;
		void OnEditAction2() override;
		void OnOK() override;
		void UpdateMethod();
		MailboxOptions& Options() noexcept { return m_options[static_cast<size_t>(m_method)]; }

		LPMAPISESSION m_session; // Borrowed for the lifetime of the modal dialog.
		MailboxMethod m_method{MailboxMethod::SMTP};
		std::array<MailboxOptions, 3> m_options{};
	};
} // namespace dialog::editor
