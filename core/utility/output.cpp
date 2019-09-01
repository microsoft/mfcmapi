#include <core/stdafx.h>
#include <core/utility/output.h>
#include <core/utility/strings.h>
#include <core/utility/file.h>
#include <core/utility/registry.h>
#include <core/utility/error.h>
#include <mapistub/library/mapiStubUtils.h>

#ifdef CHECKFORMATPARAMS
#undef Outputf
#undef OutputToFilef
#undef DebugPrint
#undef DebugPrintEx
#endif

namespace output
{
	std::wstring g_szXMLHeader = L"<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>\n";
	std::function<void(const std::wstring& errString)> outputToDbgView;
	FILE* g_fDebugFile = nullptr;

	void OpenDebugFile()
	{
		// close out the old file handle if we have one
		CloseDebugFile();

		// only open the file if we really need to
		if (registry::debugToFile)
		{
			g_fDebugFile = MyOpenFile(registry::debugFileName, false);
		}
	}

	void CloseDebugFile()
	{
		if (g_fDebugFile) CloseFile(g_fDebugFile);
		g_fDebugFile = nullptr;
	}

	// We've got our 'new' value here and also a debug output file name
	// gonna set the new value
	// gonna ensure our debug output file is open if we need it, closed if we don't
	// gonna output some text if we just toggled on
	void SetDebugOutputToFile(bool bDoOutput)
	{
		// save our old value
		const bool bOldDoOutput = registry::debugToFile;

		// set the new value
		registry::debugToFile = bDoOutput;

		// ensure we got a file if we need it
		OpenDebugFile();

		// output text if we just toggled on
		if (bDoOutput && !bOldDoOutput)
		{
			auto appName = file::GetModuleFileName(nullptr);
			DebugPrint(output::DBGGeneric, L"%ws: Debug printing to file enabled.\n", appName.c_str());

			outputVersion(output::DBGVersionBanner, nullptr);
		}
	}

	bool fIsSet(output::DBGLEVEL ulTag) { return registry::debugTag & ulTag; }
	bool fIsSetv(output::DBGLEVEL ulTag) { return (ulTag != DBGNoDebug) && (registry::debugTag & ulTag); }

	bool earlyExit(output::DBGLEVEL ulDbgLvl, bool fFile)
	{
		assert(output::DBGNoDebug != ulDbgLvl || fFile);

		// quick check to see if we have anything to print - so we can avoid executing the call
		if (!fFile && !registry::debugToFile && !fIsSetv(ulDbgLvl)) return true;
		return false;
	}

	_Check_return_ FILE* MyOpenFile(const std::wstring& szFileName, bool bNewFile)
	{
		auto mode = L"a+"; // STRING_OK
		if (bNewFile) mode = L"w"; // STRING_OK
		return MyOpenFileMode(szFileName, mode);
	}

	_Check_return_ FILE* MyOpenFileMode(const std::wstring& szFileName, const wchar_t* mode)
	{
		FILE* fOut = nullptr;

		// _wfopen has been deprecated, but older compilers do not have _wfopen_s
		// Use the replacement when we're on VS 2005 or higher.
#if defined(_MSC_VER) && (_MSC_VER >= 1400)
		_wfopen_s(&fOut, szFileName.c_str(), mode);
#else
		fOut = _wfopen(szFileName.c_str(), szParams);
#endif
		if (fOut)
		{
			return fOut;
		}

		// File IO failed - complain - not using error macros since we may not have debug output here
		const auto dwErr = GetLastError();
		auto szSysErr = strings::formatmessagesys(HRESULT_FROM_WIN32(dwErr));

		auto szErr = strings::format(
			L"_tfopen failed, hRes = 0x%08X, dwErr = 0x%08X = \"%ws\"\n", // STRING_OK
			HRESULT_FROM_WIN32(dwErr),
			dwErr,
			szSysErr.c_str());

		OutputDebugStringW(szErr.c_str());
		return nullptr;
	}

	void CloseFile(_In_opt_ FILE* fFile)
	{
		if (fFile) fclose(fFile);
	}

	void WriteFile(_In_ FILE* fFile, const std::wstring& szString)
	{
		if (!szString.empty())
		{
			auto szStringA = strings::wstringTostring(szString);
			fputs(szStringA.c_str(), fFile);
		}
	}

	void OutputThreadTime(output::DBGLEVEL ulDbgLvl)
	{
		// Compute current time and thread for a time stamp
		std::wstring szThreadTime;

		SYSTEMTIME stLocalTime = {};
		FILETIME ftLocalTime = {};

		GetSystemTime(&stLocalTime);
		GetSystemTimeAsFileTime(&ftLocalTime);

		szThreadTime = strings::format(
			L"0x%04x %02d:%02u:%02u.%03u%ws  %02u-%02u-%4u 0x%08X: ", // STRING_OK
			GetCurrentThreadId(),
			stLocalTime.wHour <= 12 ? stLocalTime.wHour : stLocalTime.wHour - 12,
			stLocalTime.wMinute,
			stLocalTime.wSecond,
			stLocalTime.wMilliseconds,
			stLocalTime.wHour <= 12 ? L"AM" : L"PM", // STRING_OK
			stLocalTime.wMonth,
			stLocalTime.wDay,
			stLocalTime.wYear,
			ulDbgLvl);
		OutputDebugStringW(szThreadTime.c_str());
		if (outputToDbgView)
		{
			outputToDbgView(szThreadTime);
		}

		// print to to our debug output log file
		if (registry::debugToFile && g_fDebugFile)
		{
			WriteFile(g_fDebugFile, szThreadTime);
		}
	}

	// The root of all debug output - call no debug output functions besides OutputDebugString from here!
	void Output(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, const std::wstring& szMsg)
	{
		if (earlyExit(ulDbgLvl, fFile)) return;
		if (szMsg.empty()) return; // nothing to print? Cool!

		// print to debug output
		if (fIsSetv(ulDbgLvl))
		{
			if (bPrintThreadTime)
			{
				OutputThreadTime(ulDbgLvl);
			}

			OutputDebugStringW(szMsg.c_str());

			if (outputToDbgView)
			{
				outputToDbgView(szMsg);
			}
			else
			{
				wprintf(L"%ws", szMsg.c_str());
			}

			// print to to our debug output log file
			if (registry::debugToFile && g_fDebugFile)
			{
				WriteFile(g_fDebugFile, szMsg);
			}
		}

		// If we were given a file - send the output there
		if (fFile)
		{
			WriteFile(fFile, szMsg);
		}
	}

	void __cdecl Outputf(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, LPCWSTR szMsg, ...)
	{
		if (earlyExit(ulDbgLvl, fFile)) return;

		va_list argList = nullptr;
		va_start(argList, szMsg);
		Output(ulDbgLvl, fFile, bPrintThreadTime, strings::formatV(szMsg, argList));
		va_end(argList);
	}

	void __cdecl OutputToFilef(_In_opt_ FILE* fFile, LPCWSTR szMsg, ...)
	{
		if (!fFile) return;

		va_list argList = nullptr;
		va_start(argList, szMsg);
		if (argList)
		{
			Output(output::DBGNoDebug, fFile, true, strings::formatV(szMsg, argList));
		}
		else
		{
			Output(output::DBGNoDebug, fFile, true, szMsg);
		}

		va_end(argList);
	}

	void __cdecl DebugPrint(output::DBGLEVEL ulDbgLvl, LPCWSTR szMsg, ...)
	{
		if (!fIsSetv(ulDbgLvl) && !registry::debugToFile) return;

		va_list argList = nullptr;
		va_start(argList, szMsg);
		DebugPrint(ulDbgLvl, szMsg, argList);
		va_end(argList);
	}

	void __cdecl DebugPrint(output::DBGLEVEL ulDbgLvl, LPCWSTR szMsg, va_list argList)
	{
		if (!fIsSetv(ulDbgLvl) && !registry::debugToFile) return;

		if (argList)
		{
			Output(ulDbgLvl, nullptr, true, strings::formatV(szMsg, argList));
		}
		else
		{
			Output(ulDbgLvl, nullptr, true, szMsg);
		}
	}

	void __cdecl DebugPrintEx(
		output::DBGLEVEL ulDbgLvl,
		std::wstring& szClass,
		const std::wstring& szFunc,
		LPCWSTR szMsg,
		...)
	{
		if (!fIsSetv(ulDbgLvl) && !registry::debugToFile) return;

		auto szMsgEx = strings::format(L"%ws::%ws %ws", szClass.c_str(), szFunc.c_str(), szMsg); // STRING_OK
		va_list argList = nullptr;
		va_start(argList, szMsg);
		if (argList)
		{
			Output(ulDbgLvl, nullptr, true, strings::formatV(szMsgEx.c_str(), argList));
		}
		else
		{
			Output(ulDbgLvl, nullptr, true, szMsgEx);
		}

		va_end(argList);
	}

	void OutputIndent(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, int iIndent)
	{
		if (earlyExit(ulDbgLvl, fFile)) return;

		for (auto i = 0; i < iIndent; i++)
			Output(ulDbgLvl, fFile, false, L"\t");
	}

#define MAXBYTES 4096
	void outputStream(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSTREAM lpStream)
	{
		if (earlyExit(ulDbgLvl, fFile)) return;

		if (!lpStream)
		{
			Output(ulDbgLvl, fFile, true, L"OutputStream called with NULL lpStream!\n");
			return;
		}

		const auto li = LARGE_INTEGER{};
		const auto hRes = WC_H_MSG(IDS_STREAMSEEKFAILED, lpStream->Seek(li, STREAM_SEEK_SET, nullptr));

		BYTE bBuf[MAXBYTES + 2]; // Allocate some extra for NULL terminators - 2 for Unicode
		ULONG ulNumBytes = 0;
		if (hRes == S_OK) do
			{
				ulNumBytes = 0;
				EC_MAPI_S(lpStream->Read(bBuf, MAXBYTES, &ulNumBytes));

				if (ulNumBytes > 0)
				{
					bBuf[ulNumBytes] = 0;
					bBuf[ulNumBytes + 1] = 0; // In case we are in Unicode
					// TODO: Check how this works in Unicode vs ANSI
					Output(
						ulDbgLvl,
						fFile,
						true,
						strings::StripCarriage(strings::LPCTSTRToWstring(reinterpret_cast<TCHAR*>(bBuf))));
				}
			} while (ulNumBytes > 0);
	}

	void outputVersion(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile)
	{
		if (earlyExit(ulDbgLvl, fFile)) return;

		// Get version information from the application.
		const auto szFullPath = file::GetModuleFileName(nullptr);
		if (!szFullPath.empty())
		{
			auto fileVersionInfo = file::GetFileVersionInfo(nullptr);

			// Load all our strings
			for (auto iVerString = IDS_VER_FIRST; iVerString <= IDS_VER_LAST; iVerString++)
			{
				const auto szVerString = strings::loadstring(iVerString);
				Outputf(
					ulDbgLvl, fFile, true, L"%ws: %ws\n", szVerString.c_str(), fileVersionInfo[szVerString].c_str());
			}
		}
	}

	void OutputCDataOpen(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile)
	{
		Output(ulDbgLvl, fFile, false, L"<![CDATA[");
	}

	void OutputCDataClose(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile) { Output(ulDbgLvl, fFile, false, L"]]>"); }

	void OutputXMLValue(
		DBGLEVEL ulDbgLvl,
		_In_opt_ FILE* fFile,
		UINT uidTag,
		const std::wstring& szValue,
		bool bWrapCData,
		int iIndent)
	{
		if (earlyExit(ulDbgLvl, fFile)) return;
		if (szValue.empty() || !uidTag) return;

		if (!szValue[0]) return;

		auto szTag = strings::loadstring(uidTag);

		OutputIndent(ulDbgLvl, fFile, iIndent);
		Outputf(ulDbgLvl, fFile, false, L"<%ws>", szTag.c_str());
		if (bWrapCData)
		{
			OutputCDataOpen(ulDbgLvl, fFile);
		}

		Output(ulDbgLvl, fFile, false, strings::ScrubStringForXML(szValue));

		if (bWrapCData)
		{
			OutputCDataClose(ulDbgLvl, fFile);
		}

		Outputf(ulDbgLvl, fFile, false, L"</%ws>\n", szTag.c_str());
	}

	void initStubCallbacks()
	{
		mapistub::logLoadMapiCallback = [](auto _1, auto _2) { output::DebugPrint(output::DBGLoadMAPI, _1, _2); };
		mapistub::logLoadLibraryCallback = [](auto _1, auto _2) { output::DebugPrint(output::DBGLoadLibrary, _1, _2); };
	}
} // namespace output