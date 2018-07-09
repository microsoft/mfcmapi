#include <StdAfx.h>
#include <UI/Dialogs/Editors/StreamEditor.h>
#include <Interpret/InterpretProp.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/String.h>
#include <UI/ViewPane/CountedTextPane.h>
#include <UI/ViewPane/SmartViewPane.h>

namespace dialog
{
	namespace editor
	{
		enum __StreamEditorTypes
		{
			EDITOR_RTF,
			EDITOR_STREAM_BINARY,
			EDITOR_STREAM_ANSI,
			EDITOR_STREAM_UNICODE,
		};

		static std::wstring CLASS = L"CStreamEditor";

		ULONG PreferredStreamType(ULONG ulPropTag)
		{
			auto ulPropType = PROP_TYPE(ulPropTag);

			if (PT_ERROR != ulPropType && PT_UNSPECIFIED != ulPropType) return ulPropTag;

			const auto ulPropID = PROP_ID(ulPropTag);

			switch (ulPropID)
			{
			case PROP_ID(PR_BODY):
				ulPropType = PT_TSTRING;
				break;
			case PROP_ID(PR_BODY_HTML):
				ulPropType = PT_TSTRING;
				break;
			case PROP_ID(PR_RTF_COMPRESSED):
				ulPropType = PT_BINARY;
				break;
			case PROP_ID(PR_ROAMING_BINARYSTREAM):
				ulPropType = PT_BINARY;
				break;
			default:
				ulPropType = PT_BINARY;
				break;
			}

			ulPropTag = CHANGE_PROP_TYPE(ulPropTag, ulPropType);
			return ulPropTag;
		}

		// Create an editor for a MAPI property - can be used to initialize stream and rtf stream editing as well
		// Takes LPMAPIPROP and ulPropTag as input - will pull SPropValue from the LPMAPIPROP
		CStreamEditor::CStreamEditor(
			_In_ CWnd* pParentWnd,
			UINT uidTitle,
			UINT uidPrompt,
			_In_ LPMAPIPROP lpMAPIProp,
			ULONG ulPropTag,
			bool bGuessType,
			bool bIsAB,
			bool bEditPropAsRTF,
			bool bUseWrapEx,
			ULONG ulRTFFlags,
			ULONG ulInCodePage,
			ULONG ulOutCodePage)
			: CEditor(pParentWnd, uidTitle, uidPrompt, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
		{
			TRACE_CONSTRUCTOR(CLASS);

			m_lpMAPIProp = lpMAPIProp;
			m_ulPropTag = ulPropTag;
			m_bAllowTypeGuessing = bGuessType;

			if (m_bAllowTypeGuessing)
			{
				m_ulPropTag = PreferredStreamType(m_ulPropTag);

				// If we're guessing and happen to be handed PR_RTF_COMPRESSED, turn on our RTF editor
				if (PR_RTF_COMPRESSED == m_ulPropTag)
				{
					bEditPropAsRTF = true;
				}
			}

			m_bIsAB = bIsAB;
			m_bUseWrapEx = bUseWrapEx;
			m_ulRTFFlags = ulRTFFlags;
			m_ulInCodePage = ulInCodePage;
			m_ulOutCodePage = ulOutCodePage;
			m_ulStreamFlags = NULL;
			m_bDocFile = false;
			m_lpStream = nullptr;
			m_StreamError = S_OK;
			m_bDisableSave = false;

			m_iTextBox = 0;
			if (bUseWrapEx)
			{
				m_iFlagBox = 1;
				m_iCodePageBox = 2;
				m_iBinBox = 3;
			}
			else
			{
				m_iFlagBox = 0xFFFFFFFF;
				m_iCodePageBox = 0xFFFFFFFF;
				m_iBinBox = 1;
			}

			m_iSmartViewBox = 0xFFFFFFFF;
			m_bDoSmartView = false;

			// Go ahead and open our property stream in case we decide to change our stream type
			// This way we can pick the right display elements
			OpenPropertyStream(false, bEditPropAsRTF);

			if (!m_bUseWrapEx && PT_BINARY == PROP_TYPE(m_ulPropTag))
			{
				m_bDoSmartView = true;
				m_iSmartViewBox = m_iBinBox + 1;
			}

			if (bEditPropAsRTF)
				m_ulEditorType = EDITOR_RTF;
			else
			{
				switch (PROP_TYPE(m_ulPropTag))
				{
				case PT_STRING8:
					m_ulEditorType = EDITOR_STREAM_ANSI;
					break;
				case PT_UNICODE:
					m_ulEditorType = EDITOR_STREAM_UNICODE;
					break;
				case PT_BINARY:
				default:
					m_ulEditorType = EDITOR_STREAM_BINARY;
					break;
				}
			}

			const auto szPromptPostFix = strings::format(
				L"\r\n%ws", interpretprop::TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, false).c_str()); // STRING_OK
			SetPromptPostFix(szPromptPostFix);

			// Let's crack our property open and see what kind of controls we'll need for it
			// One control for text stream, one for binary
			InitPane(m_iTextBox, viewpane::TextPane::CreateCollapsibleTextPane(IDS_STREAMTEXT, false));
			if (bUseWrapEx)
			{
				InitPane(m_iFlagBox, viewpane::TextPane::CreateSingleLinePane(IDS_STREAMFLAGS, true));
				InitPane(m_iCodePageBox, viewpane::TextPane::CreateSingleLinePane(IDS_CODEPAGE, true));
			}

			InitPane(m_iBinBox, viewpane::CountedTextPane::Create(IDS_STREAMBIN, false, IDS_CB));
			if (m_bDoSmartView)
			{
				InitPane(m_iSmartViewBox, viewpane::SmartViewPane::Create(IDS_SMARTVIEW));
			}
		}

		CStreamEditor::~CStreamEditor()
		{
			TRACE_DESTRUCTOR(CLASS);
			if (m_lpStream) m_lpStream->Release();
		}

		// Used to call functions which need to be called AFTER controls are created
		BOOL CStreamEditor::OnInitDialog()
		{
			const auto bRet = CEditor::OnInitDialog();

			ReadTextStreamFromProperty();

			if (m_bDoSmartView)
			{
				// Load initial smart view here
				auto lpSmartView = dynamic_cast<viewpane::SmartViewPane*>(GetPane(m_iSmartViewBox));
				if (lpSmartView)
				{
					SPropValue sProp = {0};
					sProp.ulPropTag = CHANGE_PROP_TYPE(m_ulPropTag, PT_BINARY);
					auto bin = GetBinary(m_iBinBox);
					sProp.Value.bin.lpb = bin.data();
					sProp.Value.bin.cb = ULONG(bin.size());

					// TODO: pass in named prop stuff to make this work
					const auto smartView =
						smartview::InterpretPropSmartView2(&sProp, m_lpMAPIProp, nullptr, nullptr, m_bIsAB, false);

					lpSmartView->SetParser(smartView.first);
					lpSmartView->SetStringW(smartView.second);
				}
			}

			return bRet;
		}

		void CStreamEditor::OnOK()
		{
			WriteTextStreamToProperty();
			CMyDialog::OnOK(); // don't need to call CEditor::OnOK
		}

		void CStreamEditor::OpenPropertyStream(bool bWrite, bool bRTF)
		{
			if (!m_lpMAPIProp) return;

			// Clear the previous stream if we have one
			if (m_lpStream) m_lpStream->Release();
			m_lpStream = nullptr;

			auto hRes = S_OK;
			LPSTREAM lpTmpStream = nullptr;
			auto ulRTFFlags = m_ulRTFFlags;

			output::DebugPrintEx(
				DBGStream,
				CLASS,
				L"OpenPropertyStream",
				L"opening property 0x%X (= %ws) from %p, bWrite = 0x%X\n",
				m_ulPropTag,
				interpretprop::TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, true).c_str(),
				m_lpMAPIProp,
				bWrite);

			if (bWrite)
			{
				const auto ulStgFlags = STGM_READWRITE;
				const auto ulFlags = MAPI_CREATE | MAPI_MODIFY;
				ulRTFFlags |= MAPI_MODIFY;

				if (m_bDocFile)
				{
					hRes = EC_MAPI(m_lpMAPIProp->OpenProperty(
						m_ulPropTag,
						&IID_IStreamDocfile,
						ulStgFlags,
						ulFlags,
						reinterpret_cast<LPUNKNOWN*>(&lpTmpStream)));
				}
				else
				{
					hRes = EC_MAPI(m_lpMAPIProp->OpenProperty(
						m_ulPropTag, &IID_IStream, ulStgFlags, ulFlags, reinterpret_cast<LPUNKNOWN*>(&lpTmpStream)));
				}
			}
			else
			{
				const auto ulStgFlags = STGM_READ;
				const auto ulFlags = NULL;
				hRes = WC_MAPI(m_lpMAPIProp->OpenProperty(
					m_ulPropTag, &IID_IStream, ulStgFlags, ulFlags, reinterpret_cast<LPUNKNOWN*>(&lpTmpStream)));

				// If we're guessing types, try again as a different type
				if (hRes == MAPI_E_NOT_FOUND && m_bAllowTypeGuessing)
				{
					auto ulPropTag = m_ulPropTag;
					switch (PROP_TYPE(ulPropTag))
					{
					case PT_STRING8:
					case PT_UNICODE:
						ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_BINARY);
						break;
					case PT_BINARY:
						ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_TSTRING);
						break;
					}

					if (ulPropTag != m_ulPropTag)
					{
						output::DebugPrintEx(
							DBGStream,
							CLASS,
							L"OpenPropertyStream",
							L"Retrying as 0x%X (= %ws)\n",
							m_ulPropTag,
							interpretprop::TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, true).c_str());
						hRes = WC_MAPI(m_lpMAPIProp->OpenProperty(
							ulPropTag, &IID_IStream, ulStgFlags, ulFlags, reinterpret_cast<LPUNKNOWN*>(&lpTmpStream)));
						if (SUCCEEDED(hRes))
						{
							m_ulPropTag = ulPropTag;
						}
					}
				}

				// It's possible our stream was actually an docfile - give it a try
				if (FAILED(hRes))
				{
					hRes = WC_MAPI(m_lpMAPIProp->OpenProperty(
						CHANGE_PROP_TYPE(m_ulPropTag, PT_OBJECT),
						&IID_IStreamDocfile,
						ulStgFlags,
						ulFlags,
						reinterpret_cast<LPUNKNOWN*>(&lpTmpStream)));
					if (SUCCEEDED(hRes))
					{
						m_bDocFile = true;
						m_ulPropTag = CHANGE_PROP_TYPE(m_ulPropTag, PT_OBJECT);
					}
				}
			}

			if (!bRTF)
			{
				m_lpStream = lpTmpStream;
				lpTmpStream = nullptr;
			}
			else if (lpTmpStream)
			{
				hRes = EC_H(mapi::WrapStreamForRTF(
					lpTmpStream,
					m_bUseWrapEx,
					ulRTFFlags,
					m_ulInCodePage,
					m_ulOutCodePage,
					&m_lpStream,
					&m_ulStreamFlags));
			}

			if (lpTmpStream) lpTmpStream->Release();

			// Save off any error we got to display later
			m_StreamError = hRes;
		}

		void CStreamEditor::ReadTextStreamFromProperty() const
		{
			if (!m_lpMAPIProp) return;

			output::DebugPrintEx(
				DBGStream,
				CLASS,
				L"ReadTextStreamFromProperty",
				L"opening property 0x%X (= %ws) from %p\n",
				m_ulPropTag,
				interpretprop::TagToString(m_ulPropTag, m_lpMAPIProp, m_bIsAB, true).c_str(),
				m_lpMAPIProp);

			// If we don't have a stream to display, put up an error instead
			if (FAILED(m_StreamError) || !m_lpStream)
			{
				const auto szStreamErr = strings::formatmessage(
					IDS_CANNOTOPENSTREAM, error::ErrorNameFromErrorCode(m_StreamError).c_str(), m_StreamError);
				SetStringW(m_iTextBox, szStreamErr);
				SetEditReadOnly(m_iTextBox);
				SetEditReadOnly(m_iBinBox);
				return;
			}

			if (m_bUseWrapEx) // if m_bUseWrapEx, we're read only
			{
				SetEditReadOnly(m_iTextBox);
				SetEditReadOnly(m_iBinBox);
			}

			if (m_lpStream)
			{
				auto lpPane = dynamic_cast<viewpane::TextPane*>(GetPane(m_iBinBox));
				if (lpPane)
				{
					return lpPane->SetBinaryStream(m_lpStream);
				}
			}
		}

		// this will not work if we're using WrapCompressedRTFStreamEx
		void CStreamEditor::WriteTextStreamToProperty()
		{
			// If we couldn't get a read stream, we won't be able to get a write stream
			if (!m_lpStream) return;
			if (!m_lpMAPIProp) return;
			if (!IsDirty(m_iBinBox) && !IsDirty(m_iTextBox)) return; // If we didn't change it, don't write
			if (m_bUseWrapEx) return;

			// Reopen the property stream as writeable
			OpenPropertyStream(true, EDITOR_RTF == m_ulEditorType);

			// We started with a binary stream, pull binary back into the stream
			if (m_lpStream)
			{
				// We used to use EDITSTREAM here, but instead, just use GetBinary and write it out.
				ULONG cbWritten = 0;

				auto bin = GetBinary(m_iBinBox);

				auto hRes = EC_MAPI(m_lpStream->Write(bin.data(), static_cast<ULONG>(bin.size()), &cbWritten));
				output::DebugPrintEx(DBGStream, CLASS, L"WriteTextStreamToProperty", L"wrote 0x%X\n", cbWritten);

				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI(m_lpStream->Commit(STGC_DEFAULT));
				}

				if (m_bDisableSave)
				{
					output::DebugPrintEx(DBGStream, CLASS, L"WriteTextStreamToProperty", L"Save was disabled.\n");
				}
				else
				{
					hRes = EC_MAPI(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
				}
			}

			output::DebugPrintEx(DBGStream, CLASS, L"WriteTextStreamToProperty", L"Wrote out this stream:\n");
			output::DebugPrintStream(DBGStream, m_lpStream);
		}

		_Check_return_ ULONG CStreamEditor::HandleChange(UINT nID)
		{
			const auto i = CEditor::HandleChange(nID);

			if (i == static_cast<ULONG>(-1)) return static_cast<ULONG>(-1);

			auto lpBinPane = dynamic_cast<viewpane::CountedTextPane*>(GetPane(m_iBinBox));
			if (m_iTextBox == i && lpBinPane)
			{
				switch (m_ulEditorType)
				{
				case EDITOR_STREAM_ANSI:
				case EDITOR_RTF:
				case EDITOR_STREAM_BINARY:
				default:
				{
					auto lpszA = GetStringA(m_iTextBox);

					lpBinPane->SetBinary(LPBYTE(lpszA.c_str()), lpszA.length() * sizeof(CHAR));
					lpBinPane->SetCount(lpszA.length() * sizeof(CHAR));
					break;
				}
				case EDITOR_STREAM_UNICODE:
					auto lpszW = GetStringW(m_iTextBox);

					lpBinPane->SetBinary(LPBYTE(lpszW.c_str()), lpszW.length() * sizeof(WCHAR));
					lpBinPane->SetCount(lpszW.length() * sizeof(WCHAR));
					break;
				}
			}
			else if (m_iBinBox == i)
			{
				auto bin = GetBinary(m_iBinBox);
				{
					switch (m_ulEditorType)
					{
					case EDITOR_STREAM_ANSI:
					case EDITOR_RTF:
					case EDITOR_STREAM_BINARY:
					default:
						SetStringA(m_iTextBox, std::string(LPCSTR(bin.data()), bin.size() / sizeof(CHAR)));
						if (lpBinPane) lpBinPane->SetCount(bin.size());
						break;
					case EDITOR_STREAM_UNICODE:
						SetStringW(m_iTextBox, std::wstring(LPWSTR(bin.data()), bin.size() / sizeof(WCHAR)));
						if (lpBinPane) lpBinPane->SetCount(bin.size());
						break;
					}
				}
			}

			if (m_bDoSmartView)
			{
				auto lpSmartView = dynamic_cast<viewpane::SmartViewPane*>(GetPane(m_iSmartViewBox));
				if (lpSmartView)
				{
					auto bin = GetBinary(m_iBinBox);

					SBinary Bin = {0};
					Bin.cb = ULONG(bin.size());
					Bin.lpb = bin.data();

					lpSmartView->Parse(Bin);
				}
			}

			if (m_bUseWrapEx)
			{
				auto szFlags = interpretprop::InterpretFlags(flagStreamFlag, m_ulStreamFlags);
				SetStringf(m_iFlagBox, L"0x%08X = %ws", m_ulStreamFlags, szFlags.c_str()); // STRING_OK
				SetStringW(m_iCodePageBox, strings::formatmessage(IDS_CODEPAGES, m_ulInCodePage, m_ulOutCodePage));
			}

			OnRecalcLayout();
			return i;
		}

		void CStreamEditor::SetEditReadOnly(ULONG iControl) const
		{
			auto lpPane = dynamic_cast<viewpane::TextPane*>(GetPane(iControl));
			if (lpPane)
			{
				lpPane->SetReadOnly();
			}
		}

		void CStreamEditor::DisableSave() { m_bDisableSave = true; }
	}
}