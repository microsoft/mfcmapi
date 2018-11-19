#include <StdAfx.h>
#include <UI/Dialogs/Editors/HexEditor.h>
#include <UI/FileDialogEx.h>
#include <UI/ViewPane/CountedTextPane.h>
#include <UI/ViewPane/SmartViewPane.h>
#include <UI/ViewPane/SplitterPane.h>
#include <MAPI/Cache/GlobalCache.h>
#include <MAPI/MAPIFunctions.h>

namespace dialog
{
	namespace editor
	{
		static std::wstring CLASS = L"CHexEditor";

		enum __HexEditorFields
		{
			HEXED_TEXT,
			HEXED_ANSI,
			HEXED_UNICODE,
			HEXED_BASE64,
			HEXED_HEX,
			HEXED_SMARTVIEW
		};

		CHexEditor::CHexEditor(_In_ ui::CParentWnd* pParentWnd, _In_ cache::CMapiObjects* lpMapiObjects)
			: CEditor(
				  pParentWnd,
				  IDS_HEXEDITOR,
				  NULL,
				  CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2 | CEDITOR_BUTTON_ACTION3,
				  IDS_IMPORT,
				  IDS_EXPORT,
				  IDS_CLOSE)
		{
			TRACE_CONSTRUCTOR(CLASS);
			m_lpMapiObjects = lpMapiObjects;
			if (m_lpMapiObjects) m_lpMapiObjects->AddRef();

			auto splitter = viewpane::SplitterPane::CreateHorizontalPane(HEXED_TEXT, IDS_TEXTANSIUNICODE);
			AddPane(splitter);
			splitter->SetPaneOne(viewpane::TextPane::CreateMultiLinePane(HEXED_ANSI, NULL, false));
			splitter->SetPaneTwo(viewpane::TextPane::CreateMultiLinePane(HEXED_UNICODE, NULL, false));
			AddPane(viewpane::CountedTextPane::Create(HEXED_BASE64, IDS_BASE64STRING, false, IDS_CCH));

			AddPane(viewpane::CountedTextPane::Create(HEXED_HEX, IDS_HEX, false, IDS_CB));
			auto smartViewPane = viewpane::SmartViewPane::Create(HEXED_SMARTVIEW, IDS_SMARTVIEW);
			AddPane(smartViewPane);
			DisplayParentedDialog(pParentWnd, 1000);

			smartViewPane->OnItemSelected = [&](auto _1) { return OnSmartViewNodeSelected(_1); };
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

		void CHexEditor::OnCancel() { OnOK(); }

		_Check_return_ ULONG CHexEditor::HandleChange(const UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			if (paneID == static_cast<ULONG>(-1)) return static_cast<ULONG>(-1);

			auto lpb = LPBYTE{};
			size_t cb = 0;
			std::wstring szEncodeStr;
			size_t cchEncodeStr = 0;
			switch (paneID)
			{
			case HEXED_ANSI:
			{
				auto text = GetStringA(HEXED_ANSI);
				SetStringA(HEXED_UNICODE, text);

				lpb = LPBYTE(text.c_str());
				cb = text.length() * sizeof(CHAR);

				szEncodeStr = strings::Base64Encode(cb, lpb);
				cchEncodeStr = szEncodeStr.length();
				SetStringW(HEXED_BASE64, szEncodeStr);

				SetHex(lpb, cb);
			}
			break;
			case HEXED_UNICODE: // Unicode string changed
			{
				auto text = GetStringW(HEXED_UNICODE);
				SetStringW(HEXED_ANSI, text);

				lpb = LPBYTE(text.c_str());
				cb = text.length() * sizeof(WCHAR);

				szEncodeStr = strings::Base64Encode(cb, lpb);
				cchEncodeStr = szEncodeStr.length();
				SetStringW(HEXED_BASE64, szEncodeStr);

				SetHex(lpb, cb);
			}
			break;
			case HEXED_BASE64: // base64 changed
			{
				auto szTmpString = GetStringW(HEXED_BASE64);

				// remove any whitespace before decoding
				szTmpString = strings::StripCRLF(szTmpString);

				cchEncodeStr = szTmpString.length();
				auto bin = strings::Base64Decode(szTmpString);
				lpb = bin.data();
				cb = bin.size();

				SetStringA(HEXED_ANSI, std::string(LPCSTR(lpb), cb));
				if (!(cb % 2)) // Set Unicode String
				{
					SetStringW(HEXED_UNICODE, std::wstring(LPWSTR(lpb), cb / sizeof(WCHAR)));
				}
				else
				{
					SetStringW(HEXED_UNICODE, L"");
				}

				SetHex(lpb, cb);
			}
			break;
			case HEXED_HEX: // binary changed
			{
				ClearHighlight();
				auto bin = GetBinary(HEXED_HEX);
				lpb = bin.data();
				cb = bin.size();
				SetStringA(HEXED_ANSI, std::string(LPCSTR(lpb), cb)); // ansi string

				if (!(cb % 2)) // Set Unicode String
				{
					SetStringW(HEXED_UNICODE, std::wstring(LPWSTR(lpb), cb / sizeof(WCHAR)));
				}
				else
				{
					SetStringW(HEXED_UNICODE, L"");
				}

				szEncodeStr = strings::Base64Encode(cb, lpb);
				cchEncodeStr = szEncodeStr.length();
				SetStringW(HEXED_BASE64, szEncodeStr);
			}
			break;
			default:
				break;
			}

			if (HEXED_SMARTVIEW != paneID)
			{
				// length of base64 encoded string
				auto lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(HEXED_BASE64));
				if (lpPane)
				{
					lpPane->SetCount(cchEncodeStr);
				}

				lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(HEXED_HEX));
				if (lpPane)
				{
					lpPane->SetCount(cb);
				}
			}

			// Update any parsing we've got:
			UpdateParser();

			// Force the new layout
			OnRecalcLayout();

			return paneID;
		}

		void CHexEditor::UpdateParser() const
		{
			// Find out how to interpret the data
			auto lpPane = dynamic_cast<viewpane::SmartViewPane*>(GetPane(HEXED_SMARTVIEW));
			if (lpPane)
			{
				lpPane->Parse(GetBinary(HEXED_HEX));
			}
		}

		// Import
		void CHexEditor::OnEditAction1()
		{
			auto file = file::CFileDialogExW::OpenFile(
				strings::emptystring, strings::emptystring, OFN_FILEMUSTEXIST, strings::loadstring(IDS_ALLFILES), this);
			if (!file.empty())
			{
				cache::CGlobalCache::getInstance().MAPIInitialize(NULL);
				LPSTREAM lpStream = nullptr;

				// Get a Stream interface on the input file
				EC_H_S(mapi::MyOpenStreamOnFile(MAPIAllocateBuffer, MAPIFreeBuffer, STGM_READ, file, &lpStream));

				if (lpStream)
				{
					auto lpPane = dynamic_cast<viewpane::TextPane*>(GetPane(HEXED_HEX));
					if (lpPane)
					{
						lpPane->ClearHighlight();
						lpPane->SetBinaryStream(lpStream);
					}

					lpStream->Release();
				}
			}
		}

		// Export
		void CHexEditor::OnEditAction2()
		{
			auto file = file::CFileDialogExW::SaveAs(
				strings::emptystring,
				strings::emptystring,
				OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
				strings::loadstring(IDS_ALLFILES),
				this);
			if (!file.empty())
			{
				cache::CGlobalCache::getInstance().MAPIInitialize(NULL);
				LPSTREAM lpStream = nullptr;

				// Get a Stream interface on the output file
				EC_H_S(mapi::MyOpenStreamOnFile(
					MAPIAllocateBuffer, MAPIFreeBuffer, STGM_CREATE | STGM_READWRITE, file, &lpStream));

				if (lpStream)
				{
					const auto lpPane = dynamic_cast<viewpane::TextPane*>(GetPane(HEXED_HEX));
					if (lpPane)
					{
						lpPane->GetBinaryStream(lpStream);
					}

					lpStream->Release();
				}
			}
		}

		// Close
		void CHexEditor::OnEditAction3() { OnOK(); }

		void CHexEditor::OnSmartViewNodeSelected(smartview::block* lpData)
		{
			if (!lpData) return;
			auto lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(HEXED_HEX));
			if (lpPane)
			{
				lpPane->ClearHighlight();
				auto hex = lpPane->GetStringW();
				auto start = strings::OffsetToFilteredOffset(hex, lpData->getOffset() * 2);
				auto end = strings::OffsetToFilteredOffset(hex, (lpData->getOffset() + lpData->getSize()) * 2);
				auto range = viewpane::Range{start, end};
				lpPane->AddHighlight(range);
			}
		}

		void CHexEditor::ClearHighlight()
		{
			auto lpPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(HEXED_HEX));
			if (lpPane)
			{
				lpPane->ClearHighlight();
			}
		}

		void CHexEditor::SetHex(_In_opt_count_(cb) LPBYTE lpb, size_t cb)
		{
			SetBinary(HEXED_HEX, lpb, cb);
			ClearHighlight();
		}

	} // namespace editor
} // namespace dialog