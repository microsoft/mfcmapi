#include <StdAfx.h>
#include <UI/ViewPane/TextPane.h>
#include <Interpret/String.h>
#include <UI/UIFunctions.h>

namespace viewpane
{
	static std::wstring CLASS = L"TextPane";

	TextPane* TextPane::CreateMultiLinePane(const int paneID, const UINT uidLabel, const bool bReadOnly)
	{
		return CreateSingleLinePane(paneID, uidLabel, bReadOnly, true);
	}

	TextPane* TextPane::CreateMultiLinePane(
		const int paneID,
		const UINT uidLabel,
		_In_ const std::wstring& szVal,
		const bool bReadOnly)
	{
		return CreateSingleLinePane(paneID, uidLabel, szVal, bReadOnly, true);
	}

	TextPane*
	TextPane::CreateSingleLinePane(const int paneID, const UINT uidLabel, const bool bReadOnly, const bool bMultiLine)
	{
		auto lpPane = new (std::nothrow) TextPane();
		if (lpPane)
		{
			lpPane->m_bMultiline = bMultiLine;
			lpPane->SetLabel(uidLabel, bReadOnly);
			lpPane->m_paneID = paneID;
		}

		return lpPane;
	}

	TextPane* TextPane::CreateSingleLinePane(
		const int paneID,
		const UINT uidLabel,
		_In_ const std::wstring& szVal,
		const bool bReadOnly,
		const bool bMultiLine)
	{
		auto lpPane = new (std::nothrow) TextPane();
		if (lpPane)
		{
			lpPane->m_bMultiline = bMultiLine;
			lpPane->SetLabel(uidLabel, bReadOnly);
			lpPane->SetStringW(szVal);
			lpPane->m_paneID = paneID;
		}

		return lpPane;
	}

	TextPane*
	TextPane::CreateSingleLinePaneID(const int paneID, const UINT uidLabel, const UINT uidVal, const bool bReadOnly)
	{
		auto lpPane = new (std::nothrow) TextPane();

		if (lpPane && uidVal)
		{
			lpPane->SetLabel(uidLabel, bReadOnly);
			lpPane->SetStringW(strings::loadstring(uidVal));
			lpPane->m_paneID = paneID;
		}

		return lpPane;
	}

	TextPane* TextPane::CreateCollapsibleTextPane(const int paneID, const UINT uidLabel, const bool bReadOnly)
	{
		auto pane = new (std::nothrow) TextPane();
		if (pane)
		{
			pane->SetMultiline();
			pane->SetLabel(uidLabel, bReadOnly);
			pane->m_bCollapsible = true;
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

		output::DebugPrint(DBGStream, L"EditStreamReadCallBack: cb = %d\n", cb);

		const auto cbTemp = cb / 2;
		ULONG cbTempRead = 0;
		const auto pbTempBuff = new (std::nothrow) BYTE[cbTemp];

		if (pbTempBuff)
		{
			EC_MAPI_S(stmData->Read(pbTempBuff, cbTemp, &cbTempRead));
			output::DebugPrint(DBGStream, L"EditStreamReadCallBack: read %u bytes\n", cbTempRead);

			memset(pbBuff, 0, cbTempRead * 2);
			ULONG iBinPos = 0;
			for (ULONG i = 0; i < cbTempRead; i++)
			{
				const auto bLow = static_cast<BYTE>(pbTempBuff[i] & 0xf);
				const auto bHigh = static_cast<BYTE>(pbTempBuff[i] >> 4 & 0xf);
				const auto szLow = static_cast<CHAR>(bLow <= 0x9 ? '0' + bLow : 'A' + bLow - 0xa);
				const auto szHigh = static_cast<CHAR>(bHigh <= 0x9 ? '0' + bHigh : 'A' + bHigh - 0xa);

				pbBuff[iBinPos] = szHigh;
				pbBuff[iBinPos + 1] = szLow;

				iBinPos += 2;
			}

			*pcb = cbTempRead * 2;

			delete[] pbTempBuff;
		}

		return 0;
	}

	bool TextPane::IsDirty() { return m_EditBox.m_hWnd && m_EditBox.GetModify(); }

	int TextPane::GetFixedHeight()
	{
		auto iHeight = 0;
		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		if (m_bCollapsible)
		{
			// Our expand/collapse button
			iHeight += m_iButtonHeight;
		}
		else if (!m_szLabel.empty())
		{
			// Text labels will bump directly against their edit control, so we don't add a margin here
			iHeight += m_iLabelHeight;
		}

		// A small margin between our button and the edit control, if we're collapsible and not collapsed
		if (!m_bCollapsed && m_bCollapsible)
		{
			iHeight += m_iSmallHeightMargin;
		}

		if (!m_bMultiline)
		{
			iHeight += m_iEditHeight;
		}

		iHeight += m_iSmallHeightMargin; // Bottom margin

		return iHeight;
	}

	int TextPane::GetLines()
	{
		if (m_bCollapsed)
		{
			return 0;
		}

		if (m_bMultiline)
		{
			return LINES_MULTILINEEDIT;
		}

		return 0;
	}

	void TextPane::DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height)
	{
		auto iVariableHeight = height - GetFixedHeight();
		if (0 != m_paneID)
		{
			y += m_iSmallHeightMargin;
			height -= m_iSmallHeightMargin;
		}

		const auto cmdShow = m_bCollapsed ? SW_HIDE : SW_SHOW;
		EC_B_S(m_EditBox.ShowWindow(cmdShow));
		ViewPane::DeferWindowPos(hWinPosInfo, x, y, width, height);

		if (m_bCollapsible)
		{
			y += m_iLabelHeight + m_iSmallHeightMargin;
		}
		else
		{
			if (!m_szLabel.empty())
			{
				y += m_iLabelHeight;
				height -= m_iLabelHeight;
			}

			height -= m_iSmallHeightMargin; // This is the bottom margin
		}

		EC_B_S(::DeferWindowPos(
			hWinPosInfo,
			m_EditBox.GetSafeHwnd(),
			nullptr,
			x,
			y,
			width,
			m_bCollapsible ? iVariableHeight : height,
			SWP_NOZORDER));
	}

	void TextPane::Initialize(_In_ CWnd* pParent, _In_ HDC /*hdc*/)
	{
		ViewPane::Initialize(pParent, nullptr);

		EC_B_S(m_EditBox.Create(
			WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_BORDER | WS_VISIBLE | WS_VSCROLL | ES_AUTOVSCROLL |
				(m_bReadOnly ? ES_READONLY : 0) | (m_bMultiline ? (ES_MULTILINE | ES_WANTRETURN) : (ES_AUTOHSCROLL)),
			CRect(0, 0, 0, 0),
			pParent,
			m_nID));
		ui::SubclassEdit(m_EditBox.m_hWnd, pParent ? pParent->m_hWnd : nullptr, m_bReadOnly);
		SendMessage(m_EditBox.m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(ui::GetSegoeFont()), false);

		m_bInitialized = true; // We can now call SetEditBoxText

		// Set maximum text size
		// Use -1 to allow for VERY LARGE strings
		(void) ::SendMessage(m_EditBox.m_hWnd, EM_EXLIMITTEXT, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));

		SetEditBoxText();

		m_EditBox.SetEventMask(ENM_CHANGE);

		// Clear the modify bits so we can detect changes later
		m_EditBox.SetModify(false);

		// Remove the awful autoselect of the edit control that scrolls to the end of multiline text
		if (m_bMultiline)
		{
			::PostMessage(m_EditBox.m_hWnd, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(0));
		}
	}

	struct FakeStream
	{
		std::wstring lpszW;
		size_t cbszW{};
		size_t cbCur{};
	};

	_Check_return_ static DWORD CALLBACK
	FakeEditStreamReadCallBack(const DWORD_PTR dwCookie, _In_ LPBYTE pbBuff, const LONG cb, _In_count_(cb) LONG* pcb)
	{
		if (!pbBuff || !pcb || !dwCookie) return 0;

		auto lpfs = reinterpret_cast<FakeStream*>(dwCookie);
		if (!lpfs) return 0;
		const auto cbRemaining = static_cast<ULONG>(lpfs->cbszW - lpfs->cbCur);
		const auto cbRead = min((ULONG) cb, cbRemaining);

		*pcb = cbRead;

		if (cbRead) memcpy(pbBuff, LPBYTE(lpfs->lpszW.c_str()) + lpfs->cbCur, cbRead);

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
		output::DebugPrintEx(DBGStream, CLASS, L"SetEditBoxText", L"read %d bytes from the stream\n", lBytesRead);

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

	void TextPane::SetBinary(_In_opt_count_(cb) LPBYTE lpb, const size_t cb)
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
		m_EditBox.SetBackgroundColor(false, MyGetSysColor(ui::cBackgroundReadOnly));
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
		std::wstring value;
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
			auto buffer = new (std::nothrow) BYTE[cbBuffer];
			if (buffer)
			{
				GETTEXTEX getText = {};
				getText.cb = DWORD(cbBuffer);
				getText.flags = GT_DEFAULT;
				getText.codepage = 1200;

				auto cchW = ::SendMessage(
					m_EditBox.m_hWnd,
					EM_GETTEXTEX,
					reinterpret_cast<WPARAM>(&getText),
					reinterpret_cast<LPARAM>(buffer));
				if (cchW != 0)
				{
					value = std::wstring(LPWSTR(buffer), cchText);
				}
				else
				{
					// Didn't get a string from EM_GETTEXTEX, fall back to WM_GETTEXT
					cchW = ::SendMessage(
						m_EditBox.m_hWnd,
						WM_GETTEXT,
						static_cast<WPARAM>(cchTextWithNULL),
						reinterpret_cast<LPARAM>(buffer));
					if (cchW != 0)
					{
						value = strings::stringTowstring(std::string(LPSTR(buffer), cchText));
					}
				}
			}

			delete[] buffer;
		}

		return value;
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
		const UINT uFormat = SF_TEXT;

		es.dwCookie = reinterpret_cast<DWORD_PTR>(lpStreamIn);

		// read the 'text' stream into control
		const auto lBytesRead = m_EditBox.StreamIn(uFormat, es);
		output::DebugPrintEx(DBGStream, CLASS, L"InitEditFromStream", L"read %d bytes from the stream\n", lBytesRead);

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
				DBGStream, CLASS, L"WriteToBinaryStream", L"wrote 0x%X bytes to the stream\n", cbWritten);

			if (SUCCEEDED(hRes))
			{
				EC_MAPI_S(lpStreamOut->Commit(STGC_DEFAULT));
			}
		}
	}

	void TextPane::ShowWindow(const int nCmdShow) { m_EditBox.ShowWindow(nCmdShow); }

	void TextPane::DoHighlights()
	{
		// Disable events and turn off redraw so we can mess with the control without flicker and stuff
		const auto eventMask = ::SendMessage(m_EditBox.GetSafeHwnd(), EM_SETEVENTMASK, 0, 0);
		::SendMessage(m_EditBox.GetSafeHwnd(), WM_SETREDRAW, false, 0);

		// Grab the original range so we can restore it later
		auto originalRange = CHARRANGE{};
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
			auto charrange = CHARRANGE{range.start, range.start + range.length};
			::SendMessage(m_EditBox.GetSafeHwnd(), EM_EXSETSEL, 0, reinterpret_cast<LPARAM>(&charrange));

			charformat.crTextColor = MyGetSysColor(ui::cBackground);
			charformat.crBackColor = MyGetSysColor(ui::cText);
			::SendMessage(
				m_EditBox.GetSafeHwnd(), EM_SETCHARFORMAT, SCF_SELECTION, reinterpret_cast<LPARAM>(&charformat));
		}

		// Set our original selection range back
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_EXSETSEL, 0, reinterpret_cast<LPARAM>(&originalRange));
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_HIDESELECTION, false, 0);

		// Reenable redraw and events
		::SendMessage(m_EditBox.GetSafeHwnd(), WM_SETREDRAW, true, 0);
		InvalidateRect(m_EditBox.GetSafeHwnd(), nullptr, true);
		::SendMessage(m_EditBox.GetSafeHwnd(), EM_SETEVENTMASK, 0, eventMask);
	}
} // namespace viewpane