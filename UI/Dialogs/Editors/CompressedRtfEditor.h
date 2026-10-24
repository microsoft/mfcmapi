#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	class CCompressedRtfEditor : public CEditor
	{
	public:
		explicit CCompressedRtfEditor(CWnd* parent, _In_opt_ LPMAPIPROP message = nullptr);
		bool WasSaved() const noexcept { return m_saved; }

	private:
		BOOL OnInitDialog() override;
		void OnTimer(UINT_PTR timer);
		void OnDestroy();
		void OnEditAction1() override;
		void OnEditAction2() override;
		void OnEditAction3() override;
		void Convert(ULONG source);
		void UpdateNativeBody();
		void ShowResult(HRESULT hRes, UINT success);
		void StopTimer();
		void UpdateCounts();
		_Check_return_ ULONG HandleChange(UINT nID) override;

		LPMAPIPROP m_message{};
		ULONG m_source{};
		UINT_PTR m_timer{};
		bool m_ready{};
		bool m_valid{};
		bool m_saved{};
		HRESULT m_mapiResult{MAPI_E_NOT_INITIALIZED};
		bool m_useWrapEx{};
		ULONG m_flags{MAPI_NATIVE_BODY};
		ULONG m_inCodePage{};
		ULONG m_outCodePage{CP_UNICODE};
		std::vector<BYTE> m_compressed;

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog::editor
