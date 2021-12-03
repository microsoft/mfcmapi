#include <StdAfx.h>
#include <UI/ViewPane/TextPane.h>
#include <core/utility/strings.h>
#include <UI/UIFunctions.h>
#include <core/utility/output.h>

namespace viewpane
{
	static std::wstring CLASS = L"TextPane";

	std::shared_ptr<TextPane> TextPane::CreateMultiLinePane(const int paneID, const UINT uidLabel, const bool bReadOnly)
	{
		return CreateSingleLinePane(paneID, uidLabel, bReadOnly, true);
	}

	std::shared_ptr<TextPane> TextPane::CreateMultiLinePane(
		const int paneID,
		const UINT uidLabel,
		_In_ const std::wstring& szVal,
		const bool bReadOnly)
	{
		return CreateSingleLinePane(paneID, uidLabel, szVal, bReadOnly, true);
	}

	std::shared_ptr<TextPane>
	TextPane::CreateSingleLinePane(const int paneID, const UINT uidLabel, const bool bReadOnly, const bool bMultiLine)
	{
		auto pane = std::make_shared<TextPane>();
		if (pane)
		{
			pane->m_bMultiline = bMultiLine;
			pane->SetLabel(uidLabel);
			pane->ViewPane::SetReadOnly(bReadOnly);
			pane->m_paneID = paneID;
		}

		return pane;
	}

	std::shared_ptr<TextPane> TextPane::CreateSingleLinePane(
		const int paneID,
		const UINT uidLabel,
		_In_ const std::wstring& szVal,
		const bool bReadOnly,
		const bool bMultiLine)
	{
		auto pane = std::make_shared<TextPane>();

		if (pane)
		{
			pane->m_bMultiline = bMultiLine;
			pane->SetLabel(uidLabel);
			pane->ViewPane::SetReadOnly(bReadOnly);
			pane->SetStringW(szVal);
			pane->m_paneID = paneID;
		}

		return pane;
	}

	std::shared_ptr<TextPane>
	TextPane::CreateSingleLinePaneID(const int paneID, const UINT uidLabel, const UINT uidVal, const bool bReadOnly)
	{
		auto pane = std::make_shared<TextPane>();

		if (pane && uidVal)
		{
			pane->SetLabel(uidLabel);
			pane->ViewPane::SetReadOnly(bReadOnly);
			pane->SetStringW(strings::loadstring(uidVal));
			pane->m_paneID = paneID;
		}

		return pane;
	}

	std::shared_ptr<TextPane>
	TextPane::CreateCollapsibleTextPane(const int paneID, const UINT uidLabel, const bool bReadOnly)
	{
		auto pane = std::make_shared<TextPane>();
		if (pane)
		{
			pane->SetMultiline();
			pane->SetLabel(uidLabel);
			pane->ViewPane::SetReadOnly(bReadOnly);
			pane->makeCollapsible();
			pane->m_paneID = paneID;
		}

		return pane;
	}

	// Imports binary data from a stream, converting it to hex format before returning
	_Check_return_ static DWORD CALLBACK
	EditStreamReadCallBack(const DWORD_PTR dwCookie, _In_ LPBYTE pbBuff, const LONG cb, _In_count_(cb) LONG* pcb)
	{
		if (!pbBuff || !pcb || !dwCookie) return 0;

		auto stmData = reinterpret_cast<LPSTREAM>(dwCookie);

		*pcb = 0;

		output::DebugPrint(output::dbgLevel::Stream, L"EditStreamReadCallBack: cb = %d\n", cb);

		const ULONG cbTemp = cb / 2;
		ULONG cbTempRead = 0;
		auto buffer = std::vector<BYTE>(cbTemp);

		EC_MAPI_S(stmData->Read(buffer.data(), cbTemp, &cbTempRead));
		output::DebugPrint(output::dbgLevel::Stream, L"EditStreamReadCallBack: read %u bytes\n", cbTempRead);

		memset(pbBuff, 0, cbTempRead * 2);
		ULONG iBinPos = 0;
		for (ULONG i = 0; i < cbTempRead && i < cbTemp; i++)
		{
			const auto ch = buffer[i];
			const auto bLow = static_cast<BYTE>(ch & 0xf);
			const auto bHigh = static_cast<BYTE>(ch >> 4 & 0xf);
			const auto szLow = static_cast<CHAR>(bLow <= 0x9 ? '0' + bLow : 'A' + bLow - 0xa);
			const auto szHigh = static_cast<CHAR>(bHigh <= 0x9 ? '0' + bHigh : 'A' + bHigh - 0xa);

			pbBuff[iBinPos] = szHigh;
			pbBuff[iBinPos + 1] = szLow;

			iBinPos += 2;
		}

		*pcb = cbTempRead * 2;

		return 0;
	}

	bool TextPane::IsDirty() { return m_EditBox.m_hWnd && m_EditBox.GetModify(); }

	int TextPane::GetFixedHeight()
	{
		auto iHeight = 0;
		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		iHeight += GetHeaderHeight();

		if (!m_bMultiline)
		{
			iHeight += m_iEditHeight;
		}

		iHeight += m_iSmallHeightMargin; // Bottom margin

		return iHeight;
	}

	int TextPane::GetLines()
	{
		if (collapsed())
		{
			return 0;
		}

		if (m_bMultiline)
		{
			return LINES_MULTILINEEDIT;
		}

		return 0;
	}

	HDWP TextPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		auto curY = y;
		const auto labelHeight = GetHeaderHeight();
		if (0 != m_paneID)
		{
			curY += m_iSmallHeightMargin;
		}

		// Layout our label
		hWinPosInfo = EC_D(HDWP, ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y)));

		if (collapsed())
		{
			WC_B_S(m_EditBox.ShowWindow(SW_HIDE));

			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo, m_EditBox.GetSafeHwnd(), x, curY, 0, 0, L"TextPane::DeferWindowPos::editbox(collapsed)");
		}
		else
		{
			auto editHeight = height - (curY - y) - m_iSmallHeightMargin;
			if (labelHeight)
			{
				curY += labelHeight + m_iSmallHeightMargin;
				editHeight -= labelHeight + m_iSmallHeightMargin;
			}

			WC_B_S(m_EditBox.ShowWindow(SW_SHOW));

			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo, m_EditBox.GetSafeHwnd(), x, curY, width, editHeight, L"TextPane::DeferWindowPos::editbox");
		}

		return hWinPosInfo;
	}

	void TextPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		ViewPane::Initialize(pParent, hdc);

		EC_B_S(m_EditBox.Create(
			WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_BORDER | WS_VISIBLE | WS_VSCROLL | ES_AUTOVSCROLL |
				(m_bReadOnly ? ES_READONLY : 0) | (m_bMultiline ? (ES_MULTILINE | ES_WANTRETURN) : (ES_AUTOHSCROLL)),
			CRect(0, 0, 0, 0),
			pParent,
			m_nID));

		if (m_EditBox.m_hWnd)
		{
			ui::SubclassEdit(m_EditBox.m_hWnd, pParent ? pParent->m_hWnd : nullptr, m_bReadOnly);
			SendMessage(m_EditBox.m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(ui::GetSegoeFont()), false);

			m_bInitialized = true; // We can now call SetEditBoxText

			// Set maximum text size
			// Use -1 to allow for VERY LARGE strings
			static_cast<void>(::SendMessage(m_EditBox.m_hWnd, EM_EXLIMITTEXT, WPARAM{0}, LPARAM{-1}));

			SetEditBoxText();

			m_EditBox.SetEventMask(ENM_CHANGE);

			// Clear the modify bits so we can detect changes later
			m_EditBox.SetModify(false);

			if (m_bMultiline)
			{
				LONG stops = 16; // 16 dialog template units. Default is 32.
				EC_B(::SendMessage(m_EditBox.m_hWnd, EM_SETTABSTOPS, WPARAM{1}, reinterpret_cast<LPARAM>(&stops)));

				// Remove the awful autoselect of the edit control that scrolls to the end of multiline text
				::PostMessage(m_EditBox.m_hWnd, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(0));
			}
		}
	}

	struct FakeStream
	{
		std::wstring lpszW;
		size_t cbszW{};
		size_t cbCur{};
	};

	_Check_return_ static DWORD CALLBACK FakeEditStreamReadCallBack(
		const DWORD_PTR dwCookie,
		_In_ LPBYTE pbBuff,
		const LONG cb,
		_In_count_(cb) LONG* pcb) noexcept
	{
		if (!pbBuff || !pcb || !dwCookie) return 0;

		auto lpfs = reinterpret_cast<FakeStream*>(dwCookie);
		if (!lpfs) return 0;
		const auto cbRemaining = static_cast<ULONG>(lpfs->cbszW - lpfs->cbCur);
		const auto cbRead = min((ULONG) cb, cbRemaining);

		*pcb = cbRead;

		if (cbRead) memcpy(pbBuff, reinterpret_cast<LPBYTE>(lpfs->lpszW.data()) + lpfs->cbCur, cbRead);

		lpfs->cbCur += cbRead;

		return 0;
	}

	void TextPane::SetEditBoxText()
	{
		if (!m_bInitialized) return;
		if (!m_EditBox.m_hWnd) return;

		const ULONG ulEventMask = m_EditBox.GetEventMask(); // Get original mask
		m_EditBox.SetEventMask(ENM_NONE);

		// In order to support strings with embedded NULLs, we're going to stream the string in
		// We don't have to build a real stream interface - we can fake a lightweight one
		FakeStream fs;
		fs.lpszW = m_lpszW;
		fs.cbszW = m_lpszW.length() * sizeof(WCHAR);
		fs.cbCur = 0;

		EDITSTREAM es = {reinterpret_cast<DWORD_PTR>(&fs), 0, FakeEditStreamReadCallBack};

		// read the 'text stream' into control
		const auto lBytesRead = m_EditBox.StreamIn(SF_TEXT | SF_UNICODE, es);
		output::DebugPrintEx(
			output::dbgLevel::Stream, CLASS, L"SetEditBoxText", L"read %d bytes from the stream\n", lBytesRead);

		// Clear the modify bit so this stream appears untouched
		m_EditBox.SetModify(false);

		m_EditBox.SetEventMask(ulEventMask); // put original mask back
	}

	// Sets m_lpszW
	void TextPane::SetStringW(const std::wstring& szMsg)
	{
		m_lpszW = szMsg;

		SetEditBoxText();
	}

	void TextPane::SetBinary(_In_opt_count_(cb) const BYTE* lpb, const size_t cb)
	{
		if (!lpb || !cb)
		{
			SetStringW(L"");
		}
		else
		{
			SetStringW(strings::BinToHexString(lpb, cb, false));
		}
	}

	// This is used by the DbgView - don't call any debugger functions here!!!
	void TextPane::AppendString(_In_ const std::wstring& szMsg)
	{
		m_EditBox.HideSelection(false, true);

		const auto cchText = m_EditBox.GetWindowTextLength();
		m_EditBox.SetSel(cchText, cchText);
		m_EditBox.ReplaceSel(strings::wstringTotstring(szMsg).c_str());
	}

	// This is used by the DbgView - don't call any debugger functions here!!!
	void TextPane::Clear()
	{
		m_lpszW.clear();
		::SendMessage(m_EditBox.m_hWnd, WM_SETTEXT, NULL, reinterpret_cast<LPARAM>(""));
	}

	void TextPane::SetReadOnly()
	{
		m_EditBox.SetBackgroundColor(false, MyGetSysColor(ui::uiColor::BackgroundReadOnly));
		m_EditBox.SetReadOnly();
	}

	void TextPane::SetMultiline() { m_bMultiline = true; }

	std::wstring TextPane::GetStringW() const
	{
		if (m_bCommitted) return m_lpszW;
		return GetUIValue();
	}

	std::wstring TextPane::GetUIValue() const
	{
		auto getTextLength = GETTEXTLENGTHEX{GTL_PRECISE | GTL_NUMCHARS, 1200}; // 1200 for Unicode

		auto lResult = ::SendMessage(
			m_EditBox.m_hWnd, EM_GETTEXTLENGTHEX, reinterpret_cast<WPARAM>(&getTextLength), static_cast<LPARAM>(0));
		if (lResult == E_INVALIDARG)
		{
			// We didn't get a length - try another method
			lResult = ::SendMessage(m_EditBox.m_hWnd, WM_GETTEXTLENGTH, static_cast<WPARAM>(0), static_cast<LPARAM>(0));
		}

		const auto cchText = static_cast<size_t>(lResult);
		if (cchText)
		{
			// Allocate a buffer large enough for either kind of string, along with a null terminator
			const auto cchTextWithNULL = cchText + 1;
			const auto cbBuffer = cchTextWithNULL * sizeof(WCHAR);
			auto buffer = std::vector<BYTE>(cbBuffer);

			auto getText = GETTEXTEX{};
			getText.cb = static_cast<DWORD>(cbBuffer);
			getText.flags = GT_DEFAULT;
			getText.codepage = 1200;

			auto cchW = ::SendMessage(
				m_EditBox.m_hWnd,
				EM_GETTEXTEX,
				reinterpret_cast<WPARAM>(&getText),
				reinterpret_cast<LPARAM>(buffer.data()));
			if (cchW != 0)
			{
				return std::wstring(reinterpret_cast<LPWSTR>(buffer.data()), cchText);
			}
			else
			{
				// Didn't get a string from EM_GETTEXTEX, fall back to WM_GETTEXT
				cchW = ::SendMessage(
					m_EditBox.m_hWnd,
					WM_GETTEXT,
					static_cast<WPARAM>(cchTextWithNULL),
					reinterpret_cast<LPARAM>(buffer.data()));
				if (cchW != 0)
				{
					return strings::stringTowstring(std::string(reinterpret_cast<LPSTR>(buffer.data()), cchText));
				}
			}
		}

		return strings::emptystring;
	}

	// Gets string from edit box and places it in m_lpszW
	void TextPane::CommitUIValues()
	{
		m_lpszW = GetUIValue();
		m_bCommitted = true;
	}

	// Takes a binary stream and initializes an edit control with the HEX version of this stream
	void TextPane::SetBinaryStream(_In_ LPSTREAM lpStreamIn)
	{
		EDITSTREAM es = {0, 0, EditStreamReadCallBack};
		constexpr UINT uFormat = SF_TEXT;

		es.dwCookie = reinterpret_cast<DWORD_PTR>(lpStreamIn);

		// read the 'text' stream into control
		const auto lBytesRead = m_EditBox.StreamIn(uFormat, es);
		output::DebugPrintEx(
			output::dbgLevel::Stream, CLASS, L"InitEditFromStream", L"read %d bytes from the stream\n", lBytesRead);

		// Clear the modify bit so this stream appears untouched
		m_EditBox.SetModify(false);
	}

	// Writes a hex pane out to a binary stream
	void TextPane::GetBinaryStream(_In_ LPSTREAM lpStreamOut) const
	{
		auto bin = strings::HexStringToBin(GetStringW());
		if (bin.data() != nullptr)
		{
			ULONG cbWritten = 0;
			const auto hRes = EC_MAPI(lpStreamOut->Write(bin.data(), static_cast<ULONG>(bin.size()), &cbWritten));
			output::DebugPrintEx(
				output::dbgLevel::Stream,
				CLASS,
				L"WriteToBinaryStream",
				L"wrote 0x%X bytes to the stream\n",
				cbWritten);

			if (SUCCEEDED(hRes))
			{
				EC_MAPI_S(lpStreamOut->Commit(STGC_DEFAULT));
			}
		}
	}

	void TextPane::DoHighlights()
	{
		// Disable events and turn off redraw so we can mess with the control without flicker and stuff
		const auto eventMask = ::SendMessage(m_EditBox.GetSafeHwnd(), EM_SETEVENTMASK, 0, 0);
		::SendMessage(m_EditBox.GetSafeHwnd(), WM_SETREDRAW, false, 0);

		// Grab the original scroll position
		const auto originalScroll = POINT{};
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_GETSCROLLPOS, 0, reinterpret_cast<LPARAM>(&originalScroll));

		// Grab the original range so we can restore it later
		const auto originalRange = CHARRANGE{};
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_EXGETSEL, 0, reinterpret_cast<LPARAM>(&originalRange));

		// Select the entire range
		auto wholeRange = CHARRANGE{0, -1};
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_EXSETSEL, 0, reinterpret_cast<LPARAM>(&wholeRange));

		// Wipe our current formatting
		auto charformat = CHARFORMAT2A{};
		charformat.cbSize = sizeof charformat;
		charformat.dwMask = CFM_COLOR | CFM_BACKCOLOR;
		charformat.dwEffects = CFE_AUTOCOLOR | CFE_AUTOBACKCOLOR;
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_SETCHARFORMAT, SCF_SELECTION, reinterpret_cast<LPARAM>(&charformat));

		// Clear out CFE_AUTOCOLOR and CFE_AUTOBACKCOLOR so we can change color
		charformat.dwEffects = 0;
		for (const auto range : m_highlights)
		{
			if (static_cast<LONG>(range.start) == -1 || static_cast<LONG>(range.end) == -1) continue;
			auto charrange = CHARRANGE{static_cast<LONG>(range.start), static_cast<LONG>(range.end)};
			::SendMessage(m_EditBox.GetSafeHwnd(), EM_EXSETSEL, 0, reinterpret_cast<LPARAM>(&charrange));

			charformat.crTextColor = MyGetSysColor(ui::uiColor::TextHighlight);
			charformat.crBackColor = MyGetSysColor(ui::uiColor::TextHighlightBackground);
			::SendMessage(
				m_EditBox.GetSafeHwnd(), EM_SETCHARFORMAT, SCF_SELECTION, reinterpret_cast<LPARAM>(&charformat));
		}

		// Set our original selection range back
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_EXSETSEL, 0, reinterpret_cast<LPARAM>(&originalRange));
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_HIDESELECTION, false, 0);

		// Set our original scroll position
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_SETSCROLLPOS, 0, reinterpret_cast<LPARAM>(&originalScroll));

		// Reenable redraw and events
		::SendMessage(m_EditBox.GetSafeHwnd(), WM_SETREDRAW, true, 0);
		InvalidateRect(m_EditBox.GetSafeHwnd(), nullptr, true);
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_SETEVENTMASK, 0, eventMask);
	}

	bool TextPane::containsWindow(HWND hWnd) const noexcept
	{
		if (m_EditBox.GetSafeHwnd() == hWnd) return true;
		return m_Header.containsWindow(hWnd);
	}
} // namespace viewpane