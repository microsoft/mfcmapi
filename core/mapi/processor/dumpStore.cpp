
// Routines used in dumping the contents of the Exchange store
// in to a log file

#include <core/stdafx.h>
#include <core/mapi/processor/dumpStore.h>
#include <core/utility/strings.h>
#include <core/utility/file.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFile.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/cache/namedProps.h>
#include <core/utility/error.h>
#include <core/interpret/proptags.h>
#include <core/mapi/mapiMemory.h>

namespace mapi::processor
{
	void SaveFolderContentsToTXT(
		_In_ LPMDB lpMDB,
		_In_ LPMAPIFOLDER lpFolder,
		bool bRegular,
		bool bAssoc,
		bool bDescend,
		HWND hWnd)
	{
		auto szDir = file::GetDirectoryPath(hWnd);

		if (!szDir.empty())
		{
			dumpStore MyDumpStore;
			MyDumpStore.InitMDB(lpMDB);
			MyDumpStore.InitFolder(lpFolder);
			MyDumpStore.InitFolderPathRoot(szDir);
			MyDumpStore.ProcessFolders(bRegular, bAssoc, bDescend);
		}
	}

	dumpStore::dumpStore() noexcept
	{
		m_fFolderProps = nullptr;
		m_fFolderContents = nullptr;
		m_fMailboxTable = nullptr;

		m_bRetryStreamProps = true;
		m_bOutputAttachments = true;
		m_bOutputMSG = false;
		m_bOutputList = false;
		m_nOutputFileCount = 0;
	}

	dumpStore::~dumpStore()
	{
		output::DebugPrint(
			output::dbgLevel::Console,
			L"Output messages(XML files) count: %d\n",
			m_nOutputFileCount);
		if (m_fFolderProps) output::CloseFile(m_fFolderProps);
		if (m_fFolderContents) output::CloseFile(m_fFolderContents);
		if (m_fMailboxTable) output::CloseFile(m_fMailboxTable);
	}

	void dumpStore::InitMessagePath(_In_ const std::wstring& szMessageFileName)
	{
		m_szMessageFileName = szMessageFileName;
	}

	void dumpStore::InitFolderPathRoot(_In_ const std::wstring& szFolderPathRoot)
	{
		m_szFolderPathRoot = szFolderPathRoot;
		if (m_szFolderPathRoot.length() >= file::MAXMSGPATH)
		{
			output::DebugPrint(
				output::dbgLevel::Generic,
				L"InitFolderPathRoot: \"%ws\" length (%d) greater than max length (%d)\n",
				m_szFolderPathRoot.c_str(),
				m_szFolderPathRoot.length(),
				file::MAXMSGPATH);
		}
	}

	void dumpStore::InitMailboxTablePathRoot(_In_ const std::wstring& szMailboxTablePathRoot)
	{
		m_szMailboxTablePathRoot = szMailboxTablePathRoot;
	}

	void dumpStore::EnableMSG() noexcept { m_bOutputMSG = true; }

	void dumpStore::EnableList() noexcept { m_bOutputList = true; }

	void dumpStore::DisableStreamRetry() noexcept { m_bRetryStreamProps = false; }

	void dumpStore::DisableEmbeddedAttachments() noexcept { m_bOutputAttachments = false; }

	void dumpStore::BeginMailboxTableWork(_In_ const std::wstring& szExchangeServerName)
	{
		if (m_bOutputList) return;
		const auto szTableContentsFile = strings::format(
			L"%ws\\MAILBOX_TABLE.xml", // STRING_OK
			m_szMailboxTablePathRoot.c_str());
		m_fMailboxTable = output::MyOpenFile(szTableContentsFile, true);
		if (m_fMailboxTable)
		{
			output::OutputToFile(m_fMailboxTable, output::g_szXMLHeader);
			output::OutputToFilef(m_fMailboxTable, L"<mailboxtable server=\"%ws\">\n", szExchangeServerName.c_str());
		}
	}

	void dumpStore::DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ const _SRow* lpSRow, ULONG /*ulCurRow*/)
	{
		if (!lpSRow || !m_fMailboxTable) return;
		if (m_bOutputList) return;

		const auto lpEmailAddress = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_EMAIL_ADDRESS);

		const auto lpDisplayName = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_DISPLAY_NAME);

		output::OutputToFile(m_fMailboxTable, L"<mailbox prdisplayname=\"");
		if (strings::CheckStringProp(lpDisplayName, PT_TSTRING))
		{
			output::OutputToFile(m_fMailboxTable, strings::LPCTSTRToWstring(lpDisplayName->Value.LPSZ));
		}
		output::OutputToFile(m_fMailboxTable, L"\" premailaddress=\"");
		if (!strings::CheckStringProp(lpEmailAddress, PT_TSTRING))
		{
			output::OutputToFile(m_fMailboxTable, strings::LPCTSTRToWstring(lpEmailAddress->Value.LPSZ));
		}
		output::OutputToFile(m_fMailboxTable, L"\">\n");

		output::outputSRow(output::dbgLevel::NoDebug, m_fMailboxTable, lpSRow, lpMDB);

		// build a path for our store's folder output:
		if (strings::CheckStringProp(lpEmailAddress, PT_TSTRING) && strings::CheckStringProp(lpDisplayName, PT_TSTRING))
		{
			const auto szTemp = strings::SanitizeFileName(strings::LPCTSTRToWstring(lpDisplayName->Value.LPSZ));

			m_szFolderPathRoot = m_szMailboxTablePathRoot + L"\\" + szTemp;

			// suppress any error here since the folder may already exist
			WC_B_S(CreateDirectoryW(m_szFolderPathRoot.c_str(), nullptr));
		}

		output::OutputToFile(m_fMailboxTable, L"</mailbox>\n");
	}

	void dumpStore::EndMailboxTableWork()
	{
		if (m_bOutputList) return;
		if (m_fMailboxTable)
		{
			output::OutputToFile(m_fMailboxTable, L"</mailboxtable>\n");
			output::CloseFile(m_fMailboxTable);
		}

		m_fMailboxTable = nullptr;
	}

	void dumpStore::BeginStoreWork() noexcept {}

	void dumpStore::EndStoreWork() noexcept {}

	void dumpStore::BeginFolderWork()
	{
		auto hRes = S_OK;
		m_szFolderPath = m_szFolderPathRoot + m_szFolderOffset;

		// We've done all the setup we need. If we're just outputting a list, we don't need to do the rest
		if (m_bOutputList) return;

		WC_B_S(CreateDirectoryW(m_szFolderPath.c_str(), nullptr));

		// Dump the folder props to a file
		// Holds file/path name for folder props
		const auto szFolderPropsFile = m_szFolderPath + L"FOLDER_PROPS.xml"; // STRING_OK
		m_fFolderProps = output::MyOpenFile(szFolderPropsFile, true);
		if (!m_fFolderProps) return;

		output::OutputToFile(m_fFolderProps, output::g_szXMLHeader);
		output::OutputToFile(m_fFolderProps, L"<folderprops>\n");

		LPSPropValue lpAllProps = nullptr;
		ULONG cValues = 0L;

		hRes = WC_H_GETPROPS(mapi::GetPropsNULL(m_lpFolder, fMapiUnicode, &cValues, &lpAllProps));
		if (FAILED(hRes))
		{
			output::OutputToFilef(m_fFolderProps, L"<properties error=\"0x%08X\" />\n", hRes);
		}
		else if (lpAllProps)
		{
			output::OutputToFile(m_fFolderProps, L"<properties listtype=\"summary\">\n");

			output::outputProperties(
				output::dbgLevel::NoDebug, m_fFolderProps, cValues, lpAllProps, m_lpFolder, m_bRetryStreamProps);

			output::OutputToFile(m_fFolderProps, L"</properties>\n");

			MAPIFreeBuffer(lpAllProps);
		}

		output::OutputToFile(m_fFolderProps, L"<HierarchyTable>\n");

		InitializeInterestingTagArray();
	}

	void dumpStore::DoFolderPerHierarchyTableRowWork(_In_ const _SRow* lpSRow)
	{
		if (m_bOutputList) return;
		if (!m_fFolderProps || !lpSRow) return;
		output::OutputToFile(m_fFolderProps, L"<row>\n");
		output::outputSRow(output::dbgLevel::NoDebug, m_fFolderProps, lpSRow, m_lpMDB);
		output::OutputToFile(m_fFolderProps, L"</row>\n");
	}

	void dumpStore::EndFolderWork()
	{
		MAPIFreeBuffer(m_lpInterestingPropTags);
		m_lpInterestingPropTags = nullptr;
		if (m_bOutputList) return;
		if (m_fFolderProps)
		{
			output::OutputToFile(m_fFolderProps, L"</HierarchyTable>\n");
			output::OutputToFile(m_fFolderProps, L"</folderprops>\n");
			output::CloseFile(m_fFolderProps);
		}

		m_fFolderProps = nullptr;
	}

	void dumpStore::BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows)
	{
		if (m_szFolderPath.empty()) return;
		if (m_bOutputList)
		{
			wprintf(L"Subject, Message Class, Filename\n");
			return;
		}

		// Holds file/path name for contents table output
		const auto szContentsTableFile = ulFlags & MAPI_ASSOCIATED
											 ? m_szFolderPath + L"ASSOCIATED_CONTENTS_TABLE.xml"
											 : m_szFolderPath + L"CONTENTS_TABLE.xml"; // STRING_OK
		m_fFolderContents = output::MyOpenFile(szContentsTableFile, true);
		if (m_fFolderContents)
		{
			output::OutputToFile(m_fFolderContents, output::g_szXMLHeader);
			output::OutputToFilef(
				m_fFolderContents,
				L"<ContentsTable TableType=\"%ws\" messagecount=\"0x%08X\">\n",
				ulFlags & MAPI_ASSOCIATED ? L"Associated Contents" : L"Contents", // STRING_OK
				ulCountRows);
		}
	}

	// Outputs a single message's details to the screen, so as to produce a list of messages
	void OutputMessageList(_In_ const _SRow* lpSRow, _In_ const std::wstring& szFolderPath, bool bOutputMSG)
	{
		if (!lpSRow || szFolderPath.empty()) return;
		if (szFolderPath.length() >= file::MAXMSGPATH) return;

		// Get required properties from the message
		auto lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT_W);
		std::wstring szSubj;
		if (lpTemp && strings::CheckStringProp(lpTemp, PT_UNICODE))
		{
			szSubj = lpTemp->Value.lpszW;
		}
		else
		{
			lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT_A);
			if (lpTemp && strings::CheckStringProp(lpTemp, PT_STRING8))
			{
				szSubj = strings::stringTowstring(lpTemp->Value.lpszA);
			}
		}

		lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_RECORD_KEY);
		SBinary recordKey = {};
		if (lpTemp && PR_RECORD_KEY == lpTemp->ulPropTag)
		{
			recordKey = mapi::getBin(lpTemp);
		}

		const auto lpMessageClass =
			PpropFindProp(lpSRow->lpProps, lpSRow->cValues, CHANGE_PROP_TYPE(PR_MESSAGE_CLASS, PT_UNSPECIFIED));

		wprintf(L"\"%ws\"", szSubj.c_str());
		if (lpMessageClass)
		{
			if (PT_STRING8 == PROP_TYPE(lpMessageClass->ulPropTag))
			{
				wprintf(L",\"%hs\"", lpMessageClass->Value.lpszA ? lpMessageClass->Value.lpszA : "");
			}
			else if (PT_UNICODE == PROP_TYPE(lpMessageClass->ulPropTag))
			{
				wprintf(L",\"%ws\"", lpMessageClass->Value.lpszW ? lpMessageClass->Value.lpszW : L"");
			}
		}

		auto szExt = L".xml"; // STRING_OK
		if (bOutputMSG) szExt = L".msg"; // STRING_OK

		auto szFileName = file::BuildFileNameAndPath(szExt, szSubj, szFolderPath, &recordKey);
		wprintf(L",\"%ws\"\n", szFileName.c_str());
	}

	bool dumpStore::DoContentsTablePerRowWork(_In_ const _SRow* lpSRow, ULONG ulCurRow)
	{
		if (m_bOutputList)
		{
			OutputMessageList(lpSRow, m_szFolderPath, m_bOutputMSG);
			return false;
		}
		if (!m_fFolderContents || !m_lpFolder) return true;

		output::OutputToFilef(m_fFolderContents, L"<message num=\"0x%08X\">\n", ulCurRow);

		output::outputSRow(output::dbgLevel::NoDebug, m_fFolderContents, lpSRow, m_lpFolder);

		output::OutputToFile(m_fFolderContents, L"</message>\n");
		return true;
	}

	void dumpStore::EndContentsTableWork()
	{
		if (m_bOutputList) return;
		if (m_fFolderContents)
		{
			output::OutputToFile(m_fFolderContents, L"</ContentsTable>\n");
			output::CloseFile(m_fFolderContents);
		}

		m_fFolderContents = nullptr;
	}

	void OutputBody(
		_In_ FILE* fMessageProps,
		_In_ LPMESSAGE lpMessage,
		ULONG ulBodyTag,
		_In_ const std::wstring& szBodyName,
		bool bWrapEx,
		ULONG ulCPID)
	{
		LPSTREAM lpStream = nullptr;
		LPSTREAM lpRTFUncompressed = nullptr;
		LPSTREAM lpOutputStream = nullptr;
		auto bUnicode = PROP_TYPE(ulBodyTag) == PT_UNICODE;

		auto hRes = WC_MAPI(
			lpMessage->OpenProperty(ulBodyTag, &IID_IStream, STGM_READ, NULL, reinterpret_cast<LPUNKNOWN*>(&lpStream)));
		// If we get MAPI_E_NOT_FOUND on a unicode stream, try again as ansi
		if (bUnicode && hRes == MAPI_E_NOT_FOUND)
		{
			bUnicode = false;
			ulBodyTag = CHANGE_PROP_TYPE(ulBodyTag, PT_STRING8);
			hRes = WC_MAPI(lpMessage->OpenProperty(
				ulBodyTag, &IID_IStream, STGM_READ, NULL, reinterpret_cast<LPUNKNOWN*>(&lpStream)));
		}

		// The only error we suppress is MAPI_E_NOT_FOUND, so if a body type isn't in the output, it wasn't on the message
		if (hRes != MAPI_E_NOT_FOUND)
		{
			output::OutputToFilef(fMessageProps, L"<body property=\"%ws\"", szBodyName.c_str());
			if (!lpStream)
			{
				output::OutputToFilef(fMessageProps, L" error=\"0x%08X\">\n", hRes);
			}
			else
			{
				if (PR_RTF_COMPRESSED != ulBodyTag)
				{
					lpOutputStream = lpStream;
				}
				else // If we're outputting RTF, we need to wrap it first
				{
					if (bWrapEx)
					{
						ULONG ulStreamFlags = NULL;
						WC_H_S(mapi::WrapStreamForRTF(
							lpStream, true, MAPI_NATIVE_BODY, ulCPID, CP_UNICODE, &lpRTFUncompressed, &ulStreamFlags));
						bUnicode = !(ulStreamFlags && MAPI_NATIVE_BODY_TYPE_RTF);
						auto szFlags = flags::InterpretFlags(flagStreamFlag, ulStreamFlags);
						output::OutputToFilef(
							fMessageProps,
							L" ulStreamFlags = \"0x%08X\" szStreamFlags= \"%ws\"",
							ulStreamFlags,
							szFlags.c_str());
						output::OutputToFilef(
							fMessageProps, L" CodePageIn = \"%u\" CodePageOut = \"%d\"", ulCPID, CP_UNICODE);
					}
					else
					{
						WC_H_S(mapi::WrapStreamForRTF(lpStream, false, NULL, NULL, NULL, &lpRTFUncompressed, nullptr));
					}
					if (!lpRTFUncompressed || FAILED(hRes))
					{
						output::OutputToFilef(fMessageProps, L" rtfWrapError=\"0x%08X\"", hRes);
					}
					else
					{
						lpOutputStream = lpRTFUncompressed;
					}
				}

				output::OutputToFile(fMessageProps, L">\n");
				if (lpOutputStream)
				{
					output::OutputCDataOpen(output::dbgLevel::NoDebug, fMessageProps);
					output::outputStream(output::dbgLevel::NoDebug, fMessageProps, lpOutputStream, bUnicode);
					output::OutputCDataClose(output::dbgLevel::NoDebug, fMessageProps);
				}
			}

			output::OutputToFile(fMessageProps, L"</body>\n");
		}

		if (lpRTFUncompressed) lpRTFUncompressed->Release();
		if (lpStream) lpStream->Release();
	}

	void dumpStore::InitProperties(const std::vector<std::wstring>& properties) { m_properties = properties; }

	void dumpStore::InitNamedProperties(const std::vector<std::wstring>& namedProperties)
	{
		m_namedProperties = namedProperties;
	}

	static const SizedSPropTagArray(3, boringProps) = {3, {PR_SUBJECT_W, PR_ENTRYID, PR_PARENT_DISPLAY_W}};

	bool PropIsBoring(ULONG ulPropTag) noexcept
	{
		for (ULONG iBoring = 0; iBoring < boringProps.cValues; iBoring++)
		{
			if (ulPropTag == boringProps.aulPropTag[iBoring])
			{
				return true;
			}
		}

		return false;
	}

	void dumpStore::InitializeInterestingTagArray()
	{
		auto count = static_cast<ULONG>(m_properties.size() + m_namedProperties.size());
		if (!m_lpFolder || count == 0) return;

		auto lpDisplayNameW = LPSPropValue{};
		std::wstring szDisplayName;
		WC_MAPI_S(mapi::HrGetOnePropEx(m_lpFolder, PR_DISPLAY_NAME_W, fMapiUnicode, &lpDisplayNameW));
		if (lpDisplayNameW && strings::CheckStringProp(lpDisplayNameW, PT_UNICODE))
		{
			szDisplayName = lpDisplayNameW->Value.lpszW;
		}

		MAPIFreeBuffer(lpDisplayNameW);

		//wprintf(L"Filtering %ws for one of %lu interesting properties\n", szDisplayName.c_str(), count);
		count += boringProps.cValues; // We're going to add some boring properties we'll use later for display

		auto lpTag = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(count));
		if (lpTag)
		{
			// Populate the array
			lpTag->cValues = count;
			auto i = ULONG{0};
			for (ULONG iBoring = 0; iBoring < boringProps.cValues; iBoring++)
			{
				mapi::setTag(lpTag, i++) = boringProps.aulPropTag[iBoring];
			}

			// Add regular properties to tag array
			for (auto& property : m_properties)
			{
				auto ulPropTag = proptags::LookupPropName(property.c_str());
				if (ulPropTag != 0)
				{
					if (i < count) mapi::setTag(lpTag, i++) = ulPropTag;
				}

				//wprintf(L"Looking for %ws = 0x%08X\n", property.c_str(), ulPropTag);
			}

			// Add named properties to tag array
			for (auto& namedProperty : m_namedProperties)
			{
				MAPINAMEID NamedID = {};
				auto ulPropType = PT_UNSPECIFIED;
				// Check if that string is a known dispid
				const auto lpNameIDEntry = cache::GetDispIDFromName(namedProperty.c_str());

				// If we matched on a dispid name, use that for our lookup
				if (lpNameIDEntry)
				{
					NamedID.ulKind = MNID_ID;
					NamedID.Kind.lID = lpNameIDEntry->lValue;
					NamedID.lpguid = const_cast<LPGUID>(lpNameIDEntry->lpGuid);
					ulPropType = lpNameIDEntry->ulType;
				}
				else
				{
					NamedID.ulKind = MNID_STRING;
					NamedID.Kind.lID = strings::wstringToUlong(namedProperty, 16);
				}

				if (NamedID.lpguid && (MNID_ID == NamedID.ulKind && NamedID.Kind.lID ||
									   MNID_STRING == NamedID.ulKind && NamedID.Kind.lpwstrName))
				{
					const auto lpNamedPropTags = cache::GetIDsFromNames(m_lpFolder, {NamedID}, 0);
					if (lpNamedPropTags && lpNamedPropTags->cValues == 1)
					{
						auto ulPropTag = CHANGE_PROP_TYPE(mapi::getTag(lpNamedPropTags, 0), ulPropType);
						if (ulPropTag != 0)
						{
							if (i < count) mapi::setTag(lpTag, i++) = ulPropTag;
							//switch (NamedID.ulKind)
							//{
							//case MNID_ID:
							//	wprintf(
							//		L"Looking for %ws = 0x%04X @ 0x%08X\n",
							//		namedProperty.c_str(),
							//		NamedID.Kind.lID,
							//		ulPropTag);
							//	break;
							//case MNID_STRING:
							//default:
							//	wprintf(L"Looking for %ws @ 0x%08X\n", namedProperty.c_str(), ulPropTag);
							//}
						}
					}

					MAPIFreeBuffer(lpNamedPropTags);
				}
			}

			m_lpInterestingPropTags = lpTag;
		}
	}

	bool dumpStore::MessageHasInterestingProperties(_In_ LPMESSAGE lpMessage)
	{
		if (!lpMessage || !m_lpInterestingPropTags) return true;

		output::DebugPrint(
			output::dbgLevel::Generic,
			L"MessageHasInterestingProperties: Looking for any 1 of %d properties\n",
			m_lpInterestingPropTags->cValues);
		auto fInteresting = false;

		LPSPropValue lpProps = nullptr;
		ULONG cVals = 0;
		WC_H_GETPROPS_S(lpMessage->GetProps(m_lpInterestingPropTags, fMapiUnicode, &cVals, &lpProps));
		if (lpProps)
		{
			for (ULONG i = 0; i < cVals; i++)
			{
				if (lpProps[i].ulPropTag == PR_SUBJECT_W)
				{
					if (lpProps[i].Value.lpszW)
						output::DebugPrint(output::dbgLevel::Generic, L"Examining %ws\n", lpProps[i].Value.lpszW);
					continue;
				}

				if (PropIsBoring(lpProps[i].ulPropTag)) continue;

				if (PROP_TYPE(lpProps[i].ulPropTag) == PT_ERROR) continue;
				if (PROP_TYPE(lpProps[i].ulPropTag) == PT_UNSPECIFIED) continue;
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"MessageHasInterestingProperties: Found interesting property 0x%08X\n",
					lpProps[i].ulPropTag);
				fInteresting = true;
				break;
			}

			//if (fInteresting)
			//{
			//	output::DebugPrint(output::dbgLevel::Console, L"Found interesting message:\n");
			//	output::outputProperties(output::dbgLevel::Console, nullptr, cVals, lpProps, lpMessage, false);
			//}

			MAPIFreeBuffer(lpProps);
		}

		output::DebugPrint(
			output::dbgLevel::Generic,
			L"MessageHasInterestingProperties: message %ws interesting\n",
			(fInteresting ? L"was" : L"was not"));
		return fInteresting;
	}

	void InitMessageData(
		_In_ LPMESSAGE lpMessage,
		_In_ LPVOID lpParentMessageData,
		_In_ const std::wstring& szMessageFileName,
		_In_ const std::wstring& szFolderPath,
		_Deref_out_opt_ LPVOID* lpData)
	{
		if (!lpMessage || !lpData) return;

		*lpData = new (std::nothrow) MessageData;
		if (!*lpData) return;

		auto lpMsgData = static_cast<LPMESSAGEDATA>(*lpData);

		// If we've got a parent message, we're an attachment - use attachment filename logic
		if (lpParentMessageData)
		{
			// Take the parent filename/path, remove any extension, and append -Attach.xml
			// Should we append something for attachment number?

			// Copy the source string over
			lpMsgData->szFilePath = static_cast<LPMESSAGEDATA>(lpParentMessageData)->szFilePath;

			// Remove any extension
			lpMsgData->szFilePath = lpMsgData->szFilePath.substr(0, lpMsgData->szFilePath.find_last_of(L'.'));

			// Update file name and add extension
			lpMsgData->szFilePath += strings::format(
				L"-Attach%u.xml", static_cast<LPMESSAGEDATA>(lpParentMessageData)->ulCurAttNum); // STRING_OK

			output::OutputToFilef(
				static_cast<LPMESSAGEDATA>(lpParentMessageData)->fMessageProps,
				L"<embeddedmessage path=\"%ws\"/>\n",
				lpMsgData->szFilePath.c_str());
		}
		else if (!szMessageFileName.empty()) // if we've got a file name, use it
		{
			lpMsgData->szFilePath = szMessageFileName;
		}
		else
		{
			std::wstring szSubj; // BuildFileNameAndPath will substitute a subject if we don't find one
			SBinary recordKey = {};

			auto lpSubjectW = LPSPropValue{};
			WC_MAPI_S(mapi::HrGetOnePropEx(lpMessage, PR_SUBJECT_W, fMapiUnicode, &lpSubjectW));
			if (lpSubjectW && strings::CheckStringProp(lpSubjectW, PT_UNICODE))
			{
				szSubj = lpSubjectW->Value.lpszW;
			}
			else
			{
				auto lpSubjectA = LPSPropValue{};
				WC_MAPI_S(mapi::HrGetOnePropEx(lpMessage, PR_SUBJECT_A, fMapiUnicode, &lpSubjectA));
				if (lpSubjectA && strings::CheckStringProp(lpSubjectA, PT_STRING8))
				{
					szSubj = strings::stringTowstring(lpSubjectA->Value.lpszA);
				}

				MAPIFreeBuffer(lpSubjectA);
			}

			MAPIFreeBuffer(lpSubjectW);

			auto lpRecordKey = LPSPropValue{};
			WC_MAPI_S(mapi::HrGetOnePropEx(lpMessage, PR_RECORD_KEY, fMapiUnicode, &lpRecordKey));
			if (lpRecordKey && PR_RECORD_KEY == lpRecordKey->ulPropTag)
			{
				recordKey = mapi::getBin(lpRecordKey);
			}

			lpMsgData->szFilePath = file::BuildFileNameAndPath(L".xml", szSubj, szFolderPath, &recordKey); // STRING_OK
			MAPIFreeBuffer(lpRecordKey);
		}
	}

	void OutputMessageXML(_In_ LPMESSAGE lpMessage, bool bRetryStreamProps, _Deref_out_opt_ LPVOID* lpData)
	{
		if (!lpMessage || !lpData) return;

		auto lpMsgData = static_cast<LPMESSAGEDATA>(*lpData);

		LPSPropValue lpAllProps = nullptr;
		ULONG cValues = 0L;

		// Get all props, asking for UNICODE string properties
		const auto hRes = WC_H_GETPROPS(mapi::GetPropsNULL(lpMessage, MAPI_UNICODE, &cValues, &lpAllProps));
		if (hRes == MAPI_E_BAD_CHARWIDTH)
		{
			// Didn't like MAPI_UNICODE - fall back
			WC_H_GETPROPS_S(mapi::GetPropsNULL(lpMessage, NULL, &cValues, &lpAllProps));
		}

		if (!lpMsgData->szFilePath.empty())
		{
			output::DebugPrint(
				output::dbgLevel::Console,
				L"OutputMessagePropertiesToFile: Saving to \"%ws\"\n",
				lpMsgData->szFilePath.c_str());
			lpMsgData->fMessageProps = output::MyOpenFile(lpMsgData->szFilePath, true);

			if (lpMsgData->fMessageProps)
			{
				output::OutputToFile(lpMsgData->fMessageProps, output::g_szXMLHeader);
				output::OutputToFile(lpMsgData->fMessageProps, L"<message>\n");
				if (lpAllProps)
				{
					output::OutputToFile(lpMsgData->fMessageProps, L"<properties listtype=\"summary\">\n");
#define NUMPROPS 9
					static const SizedSPropTagArray(NUMPROPS, sptCols) = {
						NUMPROPS,
						{PR_MESSAGE_CLASS_W,
						 PR_SUBJECT_W,
						 PR_SENDER_ADDRTYPE_W,
						 PR_SENDER_EMAIL_ADDRESS_W,
						 PR_MESSAGE_DELIVERY_TIME,
						 PR_ENTRYID,
						 PR_SEARCH_KEY,
						 PR_RECORD_KEY,
						 PR_INTERNET_CPID},
					};

					for (const auto column : sptCols.aulPropTag)
					{
						const auto lpTemp = PpropFindProp(lpAllProps, cValues, column);
						if (lpTemp)
						{
							output::outputProperty(
								output::dbgLevel::NoDebug,
								lpMsgData->fMessageProps,
								lpTemp,
								lpMessage,
								bRetryStreamProps);
						}
					}

					output::OutputToFile(lpMsgData->fMessageProps, L"</properties>\n");
				}

				// Log Body
				OutputBody(lpMsgData->fMessageProps, lpMessage, PR_BODY_W, L"PR_BODY_W", false, NULL);
				OutputBody(lpMsgData->fMessageProps, lpMessage, PR_BODY_HTML_W, L"PR_BODY_HTML", false, NULL);
				OutputBody(lpMsgData->fMessageProps, lpMessage, PR_RTF_COMPRESSED, L"PR_RTF_COMPRESSED", false, NULL);

				ULONG ulInCodePage = CP_ACP; // picking CP_ACP as our default
				const auto lpTemp = PpropFindProp(lpAllProps, cValues, PR_INTERNET_CPID);
				if (lpTemp && PR_INTERNET_CPID == lpTemp->ulPropTag)
				{
					ulInCodePage = lpTemp->Value.l;
				}

				OutputBody(
					lpMsgData->fMessageProps,
					lpMessage,
					PR_RTF_COMPRESSED,
					L"WrapCompressedRTFEx best body",
					true,
					ulInCodePage);

				if (lpAllProps)
				{
					output::OutputToFile(lpMsgData->fMessageProps, L"<properties listtype=\"FullPropList\">\n");

					output::outputProperties(
						output::dbgLevel::NoDebug,
						lpMsgData->fMessageProps,
						cValues,
						lpAllProps,
						lpMessage,
						bRetryStreamProps);

					output::OutputToFile(lpMsgData->fMessageProps, L"</properties>\n");
				}
			}
		}

		MAPIFreeBuffer(lpAllProps);
	}

	void OutputMessageMSG(_In_ LPMESSAGE lpMessage, _In_ const std::wstring& szFolderPath)
	{
		enum
		{
			msgPR_SUBJECT_W,
			msgPR_RECORD_KEY,
			msgNUM_COLS
		};

		static const SizedSPropTagArray(msgNUM_COLS, msgProps) = {msgNUM_COLS, {PR_SUBJECT_W, PR_RECORD_KEY}};

		if (!lpMessage || szFolderPath.empty()) return;

		output::DebugPrint(
			output::dbgLevel::Generic, L"OutputMessageMSG: Saving message to \"%ws\"\n", szFolderPath.c_str());

		std::wstring szSubj;

		ULONG cProps = 0;
		LPSPropValue lpsProps = nullptr;
		SBinary recordKey = {};

		// Get required properties from the message
		EC_H_GETPROPS_S(lpMessage->GetProps(LPSPropTagArray(&msgProps), fMapiUnicode, &cProps, &lpsProps));
		if (cProps == 2 && lpsProps)
		{
			if (strings::CheckStringProp(&lpsProps[msgPR_SUBJECT_W], PT_UNICODE))
			{
				szSubj = lpsProps[msgPR_SUBJECT_W].Value.lpszW;
			}
			if (PR_RECORD_KEY == lpsProps[msgPR_RECORD_KEY].ulPropTag)
			{
				recordKey = mapi::getBin(lpsProps[msgPR_RECORD_KEY]);
			}
		}

		auto szFileName = file::BuildFileNameAndPath(L".msg", szSubj, szFolderPath, &recordKey); // STRING_OK
		if (!szFileName.empty())
		{
			output::DebugPrint(output::dbgLevel::Generic, L"Saving to = \"%ws\"\n", szFileName.c_str());

			WC_H_S(file::SaveToMSG(lpMessage, szFileName, fMapiUnicode != 0, nullptr, false));
		}
	}

	bool dumpStore::BeginMessageWork(
		_In_ LPMESSAGE lpMessage,
		_In_ LPVOID lpParentMessageData,
		_Deref_out_opt_ LPVOID* lpData)
	{
		if (lpData) *lpData = nullptr;
		if (lpParentMessageData && !m_bOutputAttachments) return false;
		if (m_bOutputList) return false;

		if (m_bOutputMSG)
		{
			OutputMessageMSG(lpMessage, m_szFolderPath);
			return false; // no more work necessary
		}

		InitMessageData(lpMessage, lpParentMessageData, m_szMessageFileName, m_szFolderPath, lpData);

		if (!MessageHasInterestingProperties(lpMessage)) return true;

		OutputMessageXML(lpMessage, m_bRetryStreamProps, lpData);
		m_nOutputFileCount++;
		return true;
	}

	bool dumpStore::BeginRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (!lpData) return false;
		if (m_bOutputMSG) return false; // When outputting message files, no recipient work is needed
		if (m_bOutputList) return false;

		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);
		if (lpMsgData && lpMsgData->fMessageProps)
		{
			output::OutputToFile(lpMsgData->fMessageProps, L"<recipients>\n");
		}

		return true;
	}

	void dumpStore::DoMessagePerRecipientWork(
		_In_ LPMESSAGE lpMessage,
		_In_ LPVOID lpData,
		_In_ const _SRow* lpSRow,
		ULONG ulCurRow)
	{
		if (!lpMessage || !lpData || !lpSRow) return;
		if (m_bOutputMSG) return; // When outputting message files, no recipient work is needed
		if (m_bOutputList) return;

		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

		if (lpMsgData && lpMsgData->fMessageProps)
		{
			output::OutputToFilef(lpMsgData->fMessageProps, L"<recipient num=\"0x%08X\">\n", ulCurRow);

			output::outputSRow(output::dbgLevel::NoDebug, lpMsgData->fMessageProps, lpSRow, lpMessage);

			output::OutputToFile(lpMsgData->fMessageProps, L"</recipient>\n");
		}
	}

	void dumpStore::EndRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (!lpData) return;
		if (m_bOutputMSG) return; // When outputting message files, no recipient work is needed
		if (m_bOutputList) return;
		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

		if (lpMsgData && lpMsgData->fMessageProps)
		{
			output::OutputToFile(lpMsgData->fMessageProps, L"</recipients>\n");
		}
	}

	bool dumpStore::BeginAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (m_bOutputMSG) return false; // When outputting message files, no attachment work is needed
		if (m_bOutputList) return false;
		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

		if (lpMsgData && lpMsgData->fMessageProps)
		{
			output::OutputToFile(lpMsgData->fMessageProps, L"<attachments>\n");
		}

		return true;
	}

	void dumpStore::DoMessagePerAttachmentWork(
		_In_ LPMESSAGE lpMessage,
		_In_ LPVOID lpData,
		_In_ const _SRow* lpSRow,
		_In_ LPATTACH lpAttach,
		ULONG ulCurRow)
	{
		if (!lpMessage || !lpData || !lpSRow) return;
		if (m_bOutputMSG) return; // When outputting message files, no attachment work is needed
		if (m_bOutputList) return;

		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

		lpMsgData->ulCurAttNum = ulCurRow; // set this so we can pull it if this is an embedded message
		if (!lpMsgData->fMessageProps) return;

		output::OutputToFilef(lpMsgData->fMessageProps, L"<attachment num=\"0x%08X\" filename=\"", ulCurRow);

		auto lpAttachName = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_ATTACH_FILENAME);

		if (!lpAttachName || !strings::CheckStringProp(lpAttachName, PT_TSTRING))
		{
			lpAttachName = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_DISPLAY_NAME);
		}

		if (lpAttachName && strings::CheckStringProp(lpAttachName, PT_TSTRING))
			output::OutputToFile(lpMsgData->fMessageProps, strings::LPCTSTRToWstring(lpAttachName->Value.LPSZ));
		else
			output::OutputToFile(lpMsgData->fMessageProps, L"PR_ATTACH_FILENAME not found");

		output::OutputToFile(lpMsgData->fMessageProps, L"\">\n");

		output::OutputToFile(lpMsgData->fMessageProps, L"\t<tableprops>\n");
		output::outputSRow(output::dbgLevel::NoDebug, lpMsgData->fMessageProps, lpSRow, lpMessage);

		output::OutputToFile(lpMsgData->fMessageProps, L"\t</tableprops>\n");

		if (lpAttach)
		{
			ULONG ulAllProps = 0;
			LPSPropValue lpAllProps = nullptr;
			// Let's get all props from the message and dump them.
			WC_H_GETPROPS_S(mapi::GetPropsNULL(lpAttach, fMapiUnicode, &ulAllProps, &lpAllProps));
			if (lpAllProps)
			{
				output::OutputToFile(lpMsgData->fMessageProps, L"\t<getprops>\n");
				output::outputProperties(
					output::dbgLevel::NoDebug,
					lpMsgData->fMessageProps,
					ulAllProps,
					lpAllProps,
					lpAttach,
					m_bRetryStreamProps);
				output::OutputToFile(lpMsgData->fMessageProps, L"\t</getprops>\n");
			}

			MAPIFreeBuffer(lpAllProps);
			lpAllProps = nullptr;
		}

		output::OutputToFile(lpMsgData->fMessageProps, L"</attachment>\n");
	}

	void dumpStore::EndAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (!lpData) return;
		if (m_bOutputMSG) return; // When outputting message files, no attachment work is needed
		if (m_bOutputList) return;
		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);
		if (lpMsgData && lpMsgData->fMessageProps)
		{
			output::OutputToFile(lpMsgData->fMessageProps, L"</attachments>\n");
		}
	}

	void dumpStore::EndMessageWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (m_bOutputMSG) return; // When outputting message files, no end message work is needed
		if (m_bOutputList) return;
		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

		if (lpMsgData)
		{
			if (lpMsgData->fMessageProps)
			{
				output::OutputToFile(lpMsgData->fMessageProps, L"</message>\n");
				output::CloseFile(lpMsgData->fMessageProps);
			}

			delete lpMsgData;
		}
	}
} // namespace mapi::processor