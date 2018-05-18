#include "stdafx.h"
#include "HexEditor.h"
#include <UI/FileDialogEx.h>
#include "ImportProcs.h"
#include <UI/ViewPane/CountedTextPane.h>
#include <UI/ViewPane/SmartViewPane.h>
#include <MAPI/GlobalCache.h>

static wstring CLASS = L"CHexEditor";

enum __HexEditorFields
{
	HEXED_ANSI,
	HEXED_UNICODE,
	HEXED_BASE64,
	HEXED_HEX,
	HEXED_SMARTVIEW
};

CHexEditor::CHexEditor(_In_ CParentWnd* pParentWnd, _In_ CMapiObjects* lpMapiObjects) :
	CEditor(pParentWnd, IDS_HEXEDITOR, NULL, CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2 | CEDITOR_BUTTON_ACTION3, IDS_IMPORT, IDS_EXPORT, IDS_CLOSE)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpMapiObjects = lpMapiObjects;
	if (m_lpMapiObjects) m_lpMapiObjects->AddRef();

	InitPane(HEXED_ANSI, TextPane::CreateCollapsibleTextPane(IDS_ANSISTRING, false));
	InitPane(HEXED_UNICODE, TextPane::CreateCollapsibleTextPane(IDS_UNISTRING, false));
	InitPane(HEXED_BASE64, CountedTextPane::Create(IDS_BASE64STRING, false, IDS_CCH));
	InitPane(HEXED_HEX, CountedTextPane::Create(IDS_HEX, false, IDS_CB));
	InitPane(HEXED_SMARTVIEW, SmartViewPane::Create(IDS_SMARTVIEW));
	DisplayParentedDialog(pParentWnd, 1000);
}

CHexEditor::~CHexEditor()
{
	TRACE_DESTRUCTOR(CLASS);

	if (m_lpMapiObjects) m_lpMapiObjects->Release();
}

void CHexEditor::OnOK()
{
	ShowWindow(SW_HIDE);
	delete this;
}

void CHexEditor::OnCancel()
{
	OnOK();
}

_Check_return_ ULONG CHexEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	if (static_cast<ULONG>(-1) == i) return static_cast<ULONG>(-1);

	LPBYTE lpb = nullptr;
	size_t cb = 0;
	wstring szEncodeStr;
	size_t cchEncodeStr = 0;
	switch (i)
	{
	case HEXED_ANSI:
	{
		auto text = GetStringA(HEXED_ANSI);
		SetStringA(HEXED_UNICODE, text);

		lpb = LPBYTE(text.c_str());
		cb = text.length() * sizeof(CHAR);

		szEncodeStr = Base64Encode(cb, lpb);
		cchEncodeStr = szEncodeStr.length();
		SetStringW(HEXED_BASE64, szEncodeStr);

		SetBinary(HEXED_HEX, lpb, cb);
	}
	break;
	case HEXED_UNICODE: // Unicode string changed
	{
		auto text = GetStringW(HEXED_UNICODE);
		SetStringW(HEXED_ANSI, text);

		lpb = LPBYTE(text.c_str());
		cb = text.length() * sizeof(WCHAR);

		szEncodeStr = Base64Encode(cb, lpb);
		cchEncodeStr = szEncodeStr.length();
		SetStringW(HEXED_BASE64, szEncodeStr);

		SetBinary(HEXED_HEX, lpb, cb);
	}
	break;
	case HEXED_BASE64: // base64 changed
	{
		auto szTmpString = GetStringW(HEXED_BASE64);

		// remove any whitespace before decoding
		szTmpString = strings::CleanString(szTmpString);

		cchEncodeStr = szTmpString.length();
		auto bin = Base64Decode(szTmpString);
		lpb = bin.data();
		cb = bin.size();

		SetStringA(HEXED_ANSI, string(LPCSTR(lpb), cb));
		if (!(cb % 2)) // Set Unicode String
		{
			SetStringW(HEXED_UNICODE, wstring(LPWSTR(lpb), cb / sizeof(WCHAR)));
		}
		else
		{
			SetStringW(HEXED_UNICODE, L"");
		}

		SetBinary(HEXED_HEX, lpb, cb);
	}
	break;
	case HEXED_HEX: // binary changed
	{
		auto bin = GetBinary(HEXED_HEX);
		lpb = bin.data();
		cb = bin.size();
		SetStringA(HEXED_ANSI, string(LPCSTR(lpb), cb)); // ansi string

		if (!(cb % 2)) // Set Unicode String
		{
			SetStringW(HEXED_UNICODE, wstring(LPWSTR(lpb), cb / sizeof(WCHAR)));
		}
		else
		{
			SetStringW(HEXED_UNICODE, L"");
		}

		szEncodeStr = Base64Encode(cb, lpb);
		cchEncodeStr = szEncodeStr.length();
		SetStringW(HEXED_BASE64, szEncodeStr);
	}
	break;
	default:
		break;
	}

	if (HEXED_SMARTVIEW != i)
	{
		// length of base64 encoded string
		auto lpPane = static_cast<CountedTextPane*>(GetPane(HEXED_BASE64));
		if (lpPane)
		{
			lpPane->SetCount(cchEncodeStr);
		}

		lpPane = static_cast<CountedTextPane*>(GetPane(HEXED_HEX));
		if (lpPane)
		{
			lpPane->SetCount(cb);
		}
	}

	// Update any parsing we've got:
	UpdateParser();

	// Force the new layout
	OnRecalcLayout();

	return i;
}

void CHexEditor::UpdateParser() const
{
	// Find out how to interpret the data
	auto lpPane = static_cast<SmartViewPane*>(GetPane(HEXED_SMARTVIEW));
	if (lpPane)
	{
		auto bin = GetBinary(HEXED_HEX);
		SBinary Bin = { 0 };
		Bin.lpb = bin.data();
		Bin.cb = ULONG(bin.size());
		lpPane->Parse(Bin);
	}
}

// Import
void CHexEditor::OnEditAction1()
{
	auto hRes = S_OK;
	if (S_OK == hRes)
	{
		auto file = CFileDialogExW::OpenFile(
			strings::emptystring,
			strings::emptystring,
			OFN_FILEMUSTEXIST,
			strings::loadstring(IDS_ALLFILES),
			this);
		if (!file.empty())
		{
			CGlobalCache::getInstance().MAPIInitialize(NULL);
			LPSTREAM lpStream = nullptr;

			// Get a Stream interface on the input file
			EC_H(MyOpenStreamOnFile(
				MAPIAllocateBuffer,
				MAPIFreeBuffer,
				STGM_READ,
				file,
				&lpStream));

			if (lpStream)
			{
				auto lpPane = static_cast<TextPane*>(GetPane(HEXED_HEX));
				if (lpPane)
				{
					lpPane->SetBinaryStream(lpStream);
				}

				lpStream->Release();
			}
		}
	}
}

// Export
void CHexEditor::OnEditAction2()
{
	auto hRes = S_OK;
	if (S_OK == hRes)
	{
		auto file = CFileDialogExW::SaveAs(
			strings::emptystring,
			strings::emptystring,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			strings::loadstring(IDS_ALLFILES),
			this);
		if (!file.empty())
		{
			CGlobalCache::getInstance().MAPIInitialize(NULL);
			LPSTREAM lpStream = nullptr;

			// Get a Stream interface on the output file
			EC_H(MyOpenStreamOnFile(
				MAPIAllocateBuffer,
				MAPIFreeBuffer,
				STGM_CREATE | STGM_READWRITE,
				file,
				&lpStream));

			if (lpStream)
			{
				auto lpPane = static_cast<TextPane*>(GetPane(HEXED_HEX));
				if (lpPane)
				{
					lpPane->GetBinaryStream(lpStream);
				}

				lpStream->Release();
			}
		}
	}
}

// Close
void CHexEditor::OnEditAction3()
{
	OnOK();
}