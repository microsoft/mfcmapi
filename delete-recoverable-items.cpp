#include <windows.h>
#include <mapix.h>
#include <mapiutil.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <iostream>
#include <vector>

// Helper: Recursively enumerate all folders in a mailbox
void EnumerateFolders(LPMAPIFOLDER lpFolder, std::vector<LPMAPIFOLDER>& folders) {
    if (!lpFolder) return;
    folders.push_back(lpFolder); // Add current folder
    LPMAPITABLE lpTable = nullptr;
    if (SUCCEEDED(lpFolder->GetHierarchyTable(0, &lpTable)) && lpTable) {
        enum { eidPR_ENTRYID, eidNUM_COLS };
        static const SizedSPropTagArray(eidNUM_COLS, eidCols) = { eidNUM_COLS, { PR_ENTRYID } };
        if (SUCCEEDED(lpTable->SetColumns((LPSPropTagArray)&eidCols, TBL_BATCH))) {
            SRowSet* pRows = nullptr;
            while (SUCCEEDED(lpTable->QueryRows(50, 0, &pRows)) && pRows && pRows->cRows) {
                for (ULONG i = 0; i < pRows->cRows; ++i) {
                    if (pRows->aRow[i].lpProps && PR_ENTRYID == pRows->aRow[i].lpProps[eidPR_ENTRYID].ulPropTag) {
                        LPMAPIFOLDER lpSubFolder = nullptr;
                        if (SUCCEEDED(lpFolder->OpenEntry(
                                pRows->aRow[i].lpProps[eidPR_ENTRYID].Value.bin.cb,
                                (LPENTRYID)pRows->aRow[i].lpProps[eidPR_ENTRYID].Value.bin.lpb,
                                nullptr, MAPI_BEST_ACCESS, nullptr, (LPUNKNOWN*)&lpSubFolder)) && lpSubFolder) {
                            EnumerateFolders(lpSubFolder, folders);
                            lpSubFolder->Release();
                        }
                    }
                }
                FreeProws(pRows);
                pRows = nullptr;
            }
        }
        lpTable->Release();
    }
}

int wmain(int argc, wchar_t* argv[]) {
    if (argc != 2) {
        std::wcerr << L"Usage: delete-recoverable-items user@example.com" << std::endl;
        return 1;
    }
    std::wstring smtpAddress = argv[1];
    HRESULT hr = S_OK;
    MAPIINIT_0 mapiInit = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
    hr = MAPIInitialize(&mapiInit);
    if (FAILED(hr)) {
        std::wcerr << L"MAPIInitialize failed: 0x" << std::hex << hr << std::endl;
        return 2;
    }
    LPMAPISESSION lpSession = nullptr;
    hr = MAPILogonEx(0, nullptr, nullptr, MAPI_EXTENDED | MAPI_NO_MAIL | MAPI_UNICODE | MAPI_NEW_SESSION, &lpSession);
    if (FAILED(hr) || !lpSession) {
        std::wcerr << L"MAPILogonEx failed: 0x" << std::hex << hr << std::endl;
        MAPIUninitialize();
        return 3;
    }
    LPMDB lpAdminMDB = mapi::store::OpenDefaultMessageStore(lpSession);
    if (!lpAdminMDB) {
        std::wcerr << L"OpenDefaultMessageStore failed" << std::endl;
        lpSession->Release();
        MAPIUninitialize();
        return 4;
    }
    LPMDB lpUserMDB = mapi::store::OpenOtherUsersMailbox(
        lpSession,
        lpAdminMDB,
        L"", // server name (empty = use profile default)
        L"", // mailbox DN (empty = use SMTP)
        smtpAddress,
        OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP,
        true // Use CreateStoreEntryID2 for SMTP
    );
    if (!lpUserMDB) {
        std::wcerr << L"OpenOtherUsersMailbox failed for " << smtpAddress << std::endl;
        lpAdminMDB->Release();
        lpSession->Release();
        MAPIUninitialize();
        return 5;
    }
    // Open root folder
    LPMAPIFOLDER lpRootFolder = nullptr;
    ULONG objType = 0;
    hr = lpUserMDB->OpenEntry(0, nullptr, nullptr, MAPI_BEST_ACCESS, &objType, (LPUNKNOWN*)&lpRootFolder);
    if (FAILED(hr) || !lpRootFolder) {
        std::wcerr << L"OpenEntry (root folder) failed: 0x" << std::hex << hr << std::endl;
        lpUserMDB->Release();
        lpAdminMDB->Release();
        lpSession->Release();
        MAPIUninitialize();
        return 6;
    }
    std::vector<LPMAPIFOLDER> folders;
    EnumerateFolders(lpRootFolder, folders);
    std::wcout << L"Found " << folders.size() << L" folders. Deleting all messages (hard delete)..." << std::endl;
    int failCount = 0;
    for (size_t i = 0; i < folders.size(); ++i) {
        std::wcout << L"[" << (i+1) << L"/" << folders.size() << L"] Emptying folder... ";
        HRESULT hr1 = mapi::ManuallyEmptyFolder(folders[i], FALSE, TRUE); // normal contents
        HRESULT hr2 = mapi::ManuallyEmptyFolder(folders[i], TRUE, TRUE);  // associated contents
        if (FAILED(hr1) || FAILED(hr2)) {
            std::wcout << L"FAILED (0x" << std::hex << hr1 << L", 0x" << hr2 << L")" << std::endl;
            ++failCount;
        } else {
            std::wcout << L"OK" << std::endl;
        }
        folders[i]->Release();
    }
    lpRootFolder->Release();
    lpUserMDB->Release();
    lpAdminMDB->Release();
    lpSession->Release();
    MAPIUninitialize();
    if (failCount) {
        std::wcerr << L"Failed to empty " << failCount << L" folders." << std::endl;
        return 7;
    }
    std::wcout << L"All messages deleted (hard delete) for " << smtpAddress << std::endl;
    return 0;
} 