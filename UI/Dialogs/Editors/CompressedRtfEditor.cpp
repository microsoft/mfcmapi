#include <StdAfx.h>
#include <UI/Dialogs/Editors/CompressedRtfEditor.h>
#include <core/mapi/compressedRtf.h>
#include <core/utility/strings.h>
#include <core/addin/mfcmapi.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <UI/ViewPane/CountedTextPane.h>

namespace dialog::editor
{
	enum RtfFields
	{
		RTF_INPUT,
		RTF_STATUS,
		RTF_TEXT,
		RTF_HEX,
		RTF_NATIVE_TEXT,
		RTF_NATIVE_HEX,
		RTF_PAUSE
	};

	CCompressedRtfEditor::CCompressedRtfEditor(CWnd* parent, _In_opt_ LPMAPIPROP message)
		: CEditor(
			  parent,
			  IDS_COMPRESSEDRTFEDITOR,
			  IDS_COMPRESSEDRTFPROMPT,
			  CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2 | CEDITOR_BUTTON_ACTION3 |
				  (message ? CEDITOR_BUTTON_CANCEL : 0),
			  IDS_CONVERTRTF,
			  message ? IDS_SAVERTFTOMESSAGE : IDS_CLOSE,
			  IDS_RTFWRAPEX),
		  m_message(message)
	{
		AddPane(viewpane::CountedTextPane::Create(RTF_INPUT, IDS_COMPRESSEDRTFINPUT, false, IDS_CB));
		AddPane(viewpane::TextPane::CreateSingleLinePane(RTF_STATUS, IDS_COMPRESSEDRTFSTATUS, true));
		AddPane(viewpane::CountedTextPane::Create(RTF_TEXT, IDS_COMPRESSEDRTFTEXT, false, IDS_CCH));
		AddPane(viewpane::CountedTextPane::Create(RTF_HEX, IDS_COMPRESSEDRTFHEX, false, IDS_CB));
		AddPane(viewpane::CountedTextPane::Create(RTF_NATIVE_TEXT, IDS_RTFNATIVETEXT, true, IDS_CCH));
		AddPane(viewpane::CountedTextPane::Create(RTF_NATIVE_HEX, IDS_RTFNATIVEHEX, true, IDS_CB));
		AddPane(viewpane::CheckPane::Create(RTF_PAUSE, IDS_PAUSERTFCONVERSION, false, false));
		GetPane(RTF_NATIVE_HEX)->makeCollapsible(true);
		std::dynamic_pointer_cast<viewpane::TextPane>(GetPane(RTF_NATIVE_TEXT))->SetVisible(false);
		std::dynamic_pointer_cast<viewpane::TextPane>(GetPane(RTF_NATIVE_HEX))->SetVisible(false);
		EnableScroll();
		if (message) SetPromptPostFix(strings::loadstring(IDS_COMPRESSEDRTFMESSAGEPROMPT));
	}

	BOOL CCompressedRtfEditor::OnInitDialog()
	{
		const auto ret = CEditor::OnInitDialog();
		m_ready = true;
		m_mapiResult = WC_MAPI(MAPIInitialize(nullptr));
		if (FAILED(m_mapiResult))
		{
			ShowResult(m_mapiResult, 0);
			return ret;
		}

		LoadString(RTF_STATUS, IDS_COMPRESSEDRTFREADY);
		if (m_message)
		{
			LPSPropValue codePage{};
			const auto codePageResult = WC_H_IGNORE_RET(
				std::list<HRESULT>{MAPI_E_NOT_FOUND}, HrGetOneProp(m_message, PR_INTERNET_CPID, &codePage));
			if (SUCCEEDED(codePageResult) && codePage && PROP_TYPE(codePage->ulPropTag) == PT_LONG)
				m_inCodePage = codePage->Value.l;
			if (codePage) MAPIFreeBuffer(codePage);

			std::vector<BYTE> compressed;
			const auto hRes = mapi::ReadCompressedRtf(m_message, compressed);
			if (SUCCEEDED(hRes))
			{
				SetBinary(RTF_INPUT, compressed.data(), compressed.size());
				Convert(RTF_INPUT);
			}
			else
			{
				ShowResult(hRes, 0);
			}
		}

		UpdateCounts();
		return ret;
	}

	void CCompressedRtfEditor::StopTimer()
	{
		if (m_timer) KillTimer(m_timer);
		m_timer = 0;
	}

	BEGIN_MESSAGE_MAP(CCompressedRtfEditor, CEditor)
	ON_WM_TIMER()
	ON_WM_DESTROY()
	END_MESSAGE_MAP()

	void CCompressedRtfEditor::OnTimer(UINT_PTR timer)
	{
		if (m_timer && timer == m_timer)
		{
			Convert(m_source);
			return;
		}

		CEditor::OnTimer(timer);
	}

	void CCompressedRtfEditor::OnDestroy()
	{
		StopTimer();
		if (SUCCEEDED(m_mapiResult)) MAPIUninitialize();
		m_mapiResult = MAPI_E_NOT_INITIALIZED;
		CEditor::OnDestroy();
	}

	void CCompressedRtfEditor::ShowResult(const HRESULT hRes, const UINT success)
	{
		if (FAILED(hRes))
		{
			auto message = IDS_COMPRESSEDRTFAPIERROR;
			if (hRes == E_INVALIDARG)
				message = IDS_COMPRESSEDRTFBADHEX;
			else if (hRes == MAPI_E_CORRUPT_DATA)
				message = IDS_COMPRESSEDRTFBADDATA;
			else if (hRes == MAPI_E_TOO_BIG)
				message = IDS_COMPRESSEDRTFTOOBIG;
			SetStringW(RTF_STATUS, strings::format(strings::loadstring(message).c_str(), hRes));
		}
		else
		{
			LoadString(RTF_STATUS, success);
		}
	}

	void CCompressedRtfEditor::Convert(const ULONG source)
	{
		StopTimer();
		CWaitCursor wait;
		m_source = source;
		m_valid = false;
		m_compressed.clear();
		if (FAILED(m_mapiResult))
		{
			ShowResult(m_mapiResult, 0);
			return;
		}

		std::vector<BYTE> rtf;
		auto hRes = S_OK;
		if (source == RTF_INPUT)
		{
			m_compressed = GetBinary(RTF_INPUT);
			hRes = m_compressed.empty() ? WC_H(E_INVALIDARG) : mapi::DecompressRtf(m_compressed, rtf);
		}
		else
		{
			if (source == RTF_HEX)
			{
				const auto hex = GetStringW(RTF_HEX);
				rtf = strings::HexStringToBin(hex);
				if (rtf.empty() && hex.find_first_not_of(L" \t\r\n") != std::wstring::npos) hRes = WC_H(E_INVALIDARG);
			}
			else
			{
				const auto text = GetStringA(RTF_TEXT);
				rtf.assign(text.begin(), text.end());
			}

			if (SUCCEEDED(hRes)) hRes = mapi::CompressRtf(rtf, m_compressed);
			SetBinary(RTF_INPUT, m_compressed.data(), m_compressed.size());
		}

		m_valid = SUCCEEDED(hRes);
		if (source != RTF_TEXT)
		{
			SetStringW(RTF_TEXT, L"");
			if (m_valid && !rtf.empty())
				SetStringA(RTF_TEXT, std::string(reinterpret_cast<const char*>(rtf.data()), rtf.size()));
		}

		if (source != RTF_HEX) SetBinary(RTF_HEX, m_valid ? rtf.data() : nullptr, m_valid ? rtf.size() : 0);
		ShowResult(hRes, source == RTF_INPUT ? IDS_COMPRESSEDRTFSUCCESS : IDS_COMPRESSEDRTFCOMPRESSED);
		UpdateNativeBody();
		UpdateCounts();
	}

	void CCompressedRtfEditor::UpdateNativeBody()
	{
		SetStringW(RTF_NATIVE_TEXT, L"");
		SetBinary(RTF_NATIVE_HEX, nullptr, 0);
		if (!m_useWrapEx || !m_valid) return;

		std::vector<BYTE> native;
		ULONG streamFlags{};
		auto hRes = mapi::DecompressRtf(m_compressed, native, true, m_flags, m_inCodePage, m_outCodePage, &streamFlags);
		if (SUCCEEDED(hRes))
		{
			if (m_outCodePage == CP_UNICODE && !(streamFlags & MAPI_NATIVE_BODY_TYPE_RTF))
			{
				if (native.size() % sizeof(wchar_t))
					hRes = WC_H(MAPI_E_CORRUPT_DATA);
				else
				{
					std::wstring text(native.size() / sizeof(wchar_t), L'\0');
					if (!native.empty() &&
						memcpy_s(text.data(), text.size() * sizeof(wchar_t), native.data(), native.size()) != 0)
					{
						hRes = WC_H(E_INVALIDARG);
					}

					if (SUCCEEDED(hRes)) SetStringW(RTF_NATIVE_TEXT, text);
				}
			}
			else
				SetStringA(RTF_NATIVE_TEXT, std::string(native.begin(), native.end()));
		}

		if (FAILED(hRes))
		{
			ShowResult(hRes, 0);
			return;
		}

		SetBinary(RTF_NATIVE_HEX, native.data(), native.size());
		SetStringW(
			RTF_STATUS,
			strings::format(
				L"0x%08X = %ws; %ws",
				streamFlags,
				flags::InterpretFlags(flagStreamFlag, streamFlags).c_str(),
				strings::formatmessage(IDS_CODEPAGES, m_inCodePage, m_outCodePage).c_str()));
	}

	void CCompressedRtfEditor::OnEditAction1() { Convert(m_source); }

	void CCompressedRtfEditor::OnEditAction2()
	{
		if (!m_message)
		{
			OnOK();
			return;
		}

		// Convert even while paused so stale output can never be saved.
		Convert(m_source);
		if (m_valid)
		{
			const auto hRes = mapi::SaveCompressedRtf(m_message, m_compressed);
			if (SUCCEEDED(hRes)) m_saved = true;
			ShowResult(hRes, IDS_COMPRESSEDRTFSAVED);
		}
	}

	void CCompressedRtfEditor::OnEditAction3()
	{
		StopTimer();
		CEditor options(this, IDS_RTFWRAPEX, IDS_RTFWRAPEXPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		options.AddPane(viewpane::CheckPane::Create(0, IDS_USEWRAPEX, m_useWrapEx, false));
		options.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_WRAPEXFLAGS, false));
		options.SetHex(1, m_flags);
		options.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULINCODEPAGE, false));
		options.SetDecimal(2, m_inCodePage);
		options.AddPane(viewpane::TextPane::CreateSingleLinePane(3, IDS_ULOUTCODEPAGE, false));
		options.SetDecimal(3, m_outCodePage);
		options.AddPane(viewpane::TextPane::CreateMultiLinePane(4, IDS_RTFWRAPEXFLAGREFERENCE, true));
		options.LoadString(4, IDS_RTFWRAPEXFLAGSHELP);
		if (!options.DisplayDialog())
		{
			if (!GetCheck(RTF_PAUSE)) Convert(m_source);
			return;
		}

		m_useWrapEx = options.GetCheck(0);
		m_flags = options.GetHex(1);
		m_inCodePage = options.GetDecimal(2);
		m_outCodePage = options.GetDecimal(3);
		std::dynamic_pointer_cast<viewpane::TextPane>(GetPane(RTF_NATIVE_TEXT))->SetVisible(m_useWrapEx);
		std::dynamic_pointer_cast<viewpane::TextPane>(GetPane(RTF_NATIVE_HEX))->SetVisible(m_useWrapEx);
		OnRecalcLayout();
		Convert(m_source);
	}

	void CCompressedRtfEditor::UpdateCounts()
	{
		std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(RTF_INPUT))->SetCount(GetBinary(RTF_INPUT).size());
		std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(RTF_TEXT))->SetCount(GetStringW(RTF_TEXT).size());
		std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(RTF_HEX))->SetCount(GetBinary(RTF_HEX).size());
		std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(RTF_NATIVE_TEXT))
			->SetCount(GetStringW(RTF_NATIVE_TEXT).size());
		std::dynamic_pointer_cast<viewpane::CountedTextPane>(GetPane(RTF_NATIVE_HEX))
			->SetCount(GetBinary(RTF_NATIVE_HEX).size());
	}

	_Check_return_ ULONG CCompressedRtfEditor::HandleChange(const UINT nID)
	{
		const auto pane = CEditor::HandleChange(nID);
		if (!m_ready) return pane;
		if (pane == RTF_INPUT || pane == RTF_TEXT || pane == RTF_HEX)
		{
			m_source = pane;
			m_valid = false;
			m_compressed.clear();
			SetStringW(RTF_NATIVE_TEXT, L"");
			SetBinary(RTF_NATIVE_HEX, nullptr, 0);
			if (pane != RTF_INPUT) SetBinary(RTF_INPUT, nullptr, 0);
			if (pane != RTF_TEXT) SetStringW(RTF_TEXT, L"");
			if (pane != RTF_HEX) SetBinary(RTF_HEX, nullptr, 0);
			UpdateCounts();
		}

		if (pane == RTF_INPUT || pane == RTF_TEXT || pane == RTF_HEX || pane == RTF_PAUSE)
		{
			StopTimer();
			const auto paused = GetCheck(RTF_PAUSE);
			LoadString(RTF_STATUS, paused ? IDS_COMPRESSEDRTFPAUSED : IDS_COMPRESSEDRTFREADY);
			if (!paused)
			{
				m_timer = SetTimer(1, 300, nullptr);
				const auto hRes = WC_B(m_timer != 0);
				if (FAILED(hRes)) ShowResult(hRes, 0);
			}
		}

		return pane;
	}
} // namespace dialog::editor
