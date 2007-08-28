// MySecInfo.cpp : implementation file
//

#include "stdafx.h"
#include "Error.h"

#include "MySecInfo.h"
#include "MAPIFunctions.h"
#include "Registry.h"
#include "Editor.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* CLASS = _T("CMySecInfo");

// The following array defines the permission names for Exchange objects.
SI_ACCESS siExchangeAccessesFolder[] =
{
    { &GUID_NULL, frightsReadAny,			MAKEINTRESOURCEW(IDS_ACCESSREADANY),		SI_ACCESS_GENERAL },
	{ &GUID_NULL, rightsReadOnly,			MAKEINTRESOURCEW(IDS_ACCESSREADONLY),		SI_ACCESS_GENERAL },
    { &GUID_NULL, frightsCreate,			MAKEINTRESOURCEW(IDS_ACCESSCREATE),			SI_ACCESS_GENERAL },
    { &GUID_NULL, frightsEditOwned,			MAKEINTRESOURCEW(IDS_ACCESSEDITOWN),		SI_ACCESS_GENERAL },
    { &GUID_NULL, frightsDeleteOwned,		MAKEINTRESOURCEW(IDS_ACCESSDELETEOWN),		SI_ACCESS_CONTAINER },
    { &GUID_NULL, frightsEditAny,			MAKEINTRESOURCEW(IDS_ACCESSEDITANY),		SI_ACCESS_GENERAL },
    { &GUID_NULL, rightsReadWrite,			MAKEINTRESOURCEW(IDS_ACCESSREADWRITE),		SI_ACCESS_GENERAL },
    { &GUID_NULL, frightsDeleteAny,			MAKEINTRESOURCEW(IDS_ACCESSDELETEANY),		SI_ACCESS_GENERAL },
    { &GUID_NULL, frightsCreateSubfolder,	MAKEINTRESOURCEW(IDS_ACCESSCREATESUBFOLDER),SI_ACCESS_GENERAL },
    { &GUID_NULL, frightsOwner,				MAKEINTRESOURCEW(IDS_ACCESSOWNER),			SI_ACCESS_GENERAL },
	{ &GUID_NULL, frightsVisible,			MAKEINTRESOURCEW(IDS_ACCESSVISIBLE),		SI_ACCESS_GENERAL },
    { &GUID_NULL, frightsContact,			MAKEINTRESOURCEW(IDS_ACCESSCONTACT),		SI_ACCESS_GENERAL }
};

SI_ACCESS siExchangeAccessesMessage[] =
{
	{ &GUID_NULL, fsdrightDelete,			MAKEINTRESOURCEW(IDS_ACCESSDELETE),			SI_ACCESS_GENERAL },
	{ &GUID_NULL, fsdrightReadProperty,		MAKEINTRESOURCEW(IDS_ACCESSREADPROPERTY),	SI_ACCESS_GENERAL },
	{ &GUID_NULL, fsdrightWriteProperty,	MAKEINTRESOURCEW(IDS_ACCESSWRITEPROPERTY),	SI_ACCESS_GENERAL },
//	{ &GUID_NULL, fsdrightCreateMessage,	MAKEINTRESOURCEW(IDS_ACCESSCREATEMESSAGE),	SI_ACCESS_GENERAL },
//	{ &GUID_NULL, fsdrightSaveMessage,		MAKEINTRESOURCEW(IDS_ACCESSSAVEMESSAGE),	SI_ACCESS_GENERAL },
//	{ &GUID_NULL, fsdrightOpenMessage,		MAKEINTRESOURCEW(IDS_ACCESSOPENMESSAGE),	SI_ACCESS_GENERAL },
	{ &GUID_NULL, fsdrightWriteSD,			MAKEINTRESOURCEW(IDS_ACCESSWRITESD),		SI_ACCESS_GENERAL },
	{ &GUID_NULL, fsdrightWriteOwner,		MAKEINTRESOURCEW(IDS_ACCESSWRITEOWNER),		SI_ACCESS_GENERAL },
	{ &GUID_NULL, fsdrightReadControl,		MAKEINTRESOURCEW(IDS_ACCESSREADCONTROL),	SI_ACCESS_GENERAL },
};

SI_ACCESS siExchangeAccessesFreeBusy[] =
{
	{ &GUID_NULL, fsdrightFreeBusySimple,	MAKEINTRESOURCEW(IDS_ACCESSSIMPLEFREEBUSY),		SI_ACCESS_GENERAL },
	{ &GUID_NULL, fsdrightFreeBusyDetailed,	MAKEINTRESOURCEW(IDS_ACCESSDETAILEDFREEBUSY),	SI_ACCESS_GENERAL },
};

GENERIC_MAPPING gmFolders =
{
	frightsReadAny,
	frightsEditOwned | frightsDeleteOwned,
	rightsNone,
	rightsAll
};


GENERIC_MAPPING gmMessages =
{
	msgrightsGenericRead,
	msgrightsGenericWrite,
	msgrightsGenericExecute,
	msgrightsGenericAll
};

//Returns pointer to buffer contained in a prop
//Free with MAPIFreeBuffer
HRESULT GetBuffer(LPMAPIPROP lpMAPIProp, ULONG ulPropTag, ULONG* cbBuf, LPBYTE* lpBuf)
{
	if (!lpMAPIProp || !ulPropTag || !lpBuf) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric,_T("GetBuffer getting buffer from 0x%08X\n"),ulPropTag);

	ulPropTag = CHANGE_PROP_TYPE(ulPropTag,PT_BINARY);

	HRESULT			hRes		= S_OK;
	ULONG			cValues		= 0;
	LPSPropValue	lpPropArray	= NULL;

	SizedSPropTagArray(1, sptaBuffer) = {1,{ulPropTag}};

	if (cbBuf) *cbBuf = NULL;

	WC_H(lpMAPIProp->GetProps((LPSPropTagArray)&sptaBuffer, 0, &cValues, &lpPropArray));

	if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag) && MAPI_E_NOT_ENOUGH_MEMORY == lpPropArray->Value.err)
	{
		//need to get the data as a stream
		LPSTREAM lpStream = NULL;

		WC_H(lpMAPIProp->OpenProperty(
			ulPropTag,
			&IID_IStream,
			STGM_READ,
			0,
			(LPUNKNOWN*) &lpStream));
		if (SUCCEEDED(hRes) && lpStream)
		{
			STATSTG	StatInfo = {0};
			lpStream->Stat(&StatInfo, STATFLAG_NONAME);//find out how much space we need

			EC_H(MAPIAllocateBuffer(
				StatInfo.cbSize.LowPart,
				(LPVOID*) lpBuf));

			if (*lpBuf)
			{
				ULONG cbRead = 0;
				EC_H(lpStream->Read(*lpBuf, StatInfo.cbSize.LowPart, &cbRead));
				if (SUCCEEDED(hRes) && cbRead == StatInfo.cbSize.LowPart)
				{
					if (cbBuf) *cbBuf = cbRead;
				}
				else
				{
					MAPIFreeBuffer(*lpBuf);
					*lpBuf = NULL;
				}
			}
		}
		if (lpStream) lpStream->Release();
	}
	else if (lpPropArray && lpPropArray->ulPropTag == ulPropTag)
	{
		EC_H(MAPIAllocateBuffer(
			lpPropArray->Value.bin.cb,
			(LPVOID*) lpBuf));
		if (*lpBuf)
		{
			memcpy(*lpBuf,lpPropArray->Value.bin.lpb,lpPropArray->Value.bin.cb);
			if (cbBuf) *cbBuf = lpPropArray->Value.bin.cb;
		}
	}

	MAPIFreeBuffer(lpPropArray);
	return hRes;
}

CMySecInfo::CMySecInfo(LPMAPIPROP lpMAPIProp,
					   ULONG ulPropTag)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_cRef = 1;
	m_ulPropTag = CHANGE_PROP_TYPE(ulPropTag,PT_BINARY);//An SD must be in a binary prop
	m_lpMAPIProp = lpMAPIProp;
	m_acetype = acetypeMessage;
	m_cbHeader = 0;
	m_lpHeader = NULL;

	if (m_lpMAPIProp)
	{
		ULONG m_ulObjType = 0;
		m_lpMAPIProp->AddRef();
		m_ulObjType = GetMAPIObjectType(m_lpMAPIProp);
		switch (m_ulObjType)
		{
		case (MAPI_STORE):
		case (MAPI_ADDRBOOK):
		case (MAPI_FOLDER):
		case (MAPI_ABCONT):
			m_acetype = acetypeContainer;
			break;
		}
	}

	if (PR_FREEBUSY_NT_SECURITY_DESCRIPTOR == m_ulPropTag)
		m_acetype = acetypeFreeBusy;

	HRESULT hRes = S_OK;
	int iRet = NULL;
	// CString doesn't provide a way to extract just Unicode strings, so we do this manually
	EC_D(iRet,LoadStringW(GetModuleHandle(NULL),
		IDS_OBJECT,
		m_wszObject,
		CCHW(m_wszObject)));

};
CMySecInfo::~CMySecInfo()
{
	TRACE_DESTRUCTOR(CLASS);
	MAPIFreeBuffer(m_lpHeader);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
}

STDMETHODIMP CMySecInfo::QueryInterface (REFIID riid,
										 LPVOID * ppvObj)
{
	*ppvObj = 0;
	if (riid == IID_ISecurityInformation ||
		riid == IID_IUnknown)
	{
		*ppvObj = (ISecurityInformation *) this;
		AddRef();
		return S_OK;
	}
	else if (riid == IID_ISecurityInformation2)
	{
		*ppvObj = (ISecurityInformation2 *) this;
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
};

STDMETHODIMP_(ULONG) CMySecInfo::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	return m_cRef;
}

STDMETHODIMP_(ULONG) CMySecInfo::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount) delete this;
	return lCount;
}
STDMETHODIMP CMySecInfo::GetObjectInformation(THIS_ PSI_OBJECT_INFO pObjectInfo )
{
	DebugPrint(DBGGeneric,_T("CMySecInfo::GetObjectInformation\n"));
	HRESULT hRes = S_OK;
	BOOL bAllowEdits = false;

	HKEY hRootKey = NULL;

	WC_W32(RegOpenKeyEx(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		NULL,
		KEY_READ,
		&hRootKey));

	if (hRootKey)
	{
		ReadDWORDFromRegistry(
			hRootKey,
			_T("AllowUnsupportedSecurityEdits"),// STRING_OK
			(DWORD*) &bAllowEdits);
		EC_W32(RegCloseKey(hRootKey));
	}

	if (bAllowEdits)
	{
		pObjectInfo->dwFlags = SI_EDIT_PERMS | SI_EDIT_OWNER | SI_ADVANCED | ((acetypeContainer == m_acetype) ? SI_CONTAINER : 0);
	}
	else
	{
		pObjectInfo->dwFlags = SI_READONLY | SI_ADVANCED | ((acetypeContainer == m_acetype) ? SI_CONTAINER : 0);
	}
	pObjectInfo->pszObjectName = m_wszObject;//Object being edited
	pObjectInfo->pszServerName = NULL;//specify DC for lookups
	return S_OK;
};

STDMETHODIMP CMySecInfo::GetSecurity(THIS_ SECURITY_INFORMATION /*RequestedInformation*/,
									PSECURITY_DESCRIPTOR *ppSecurityDescriptor,
									BOOL /*fDefault*/)
{
	DebugPrint(DBGGeneric,_T("CMySecInfo::GetSecurity\n"));
	HRESULT	hRes = S_OK;
	LPBYTE	lpSDBuffer = NULL;
	ULONG	cbSBBuffer = NULL;

	*ppSecurityDescriptor = NULL;

	PSECURITY_DESCRIPTOR pSecDesc = NULL;//will be a pointer into lpPropArray, do not free!

	EC_H(GetBuffer(m_lpMAPIProp,m_ulPropTag,&cbSBBuffer, &lpSDBuffer));

	if (lpSDBuffer)
	{
		pSecDesc = SECURITY_DESCRIPTOR_OF(lpSDBuffer);

		if (IsValidSecurityDescriptor(pSecDesc))
		{
			if (FCheckSecurityDescriptorVersion(lpSDBuffer))
			{
				int cbBuffer = GetSecurityDescriptorLength(pSecDesc);
				*ppSecurityDescriptor = LocalAlloc(LMEM_FIXED, cbBuffer);

				if ((*ppSecurityDescriptor))
				{
					memcpy(*ppSecurityDescriptor, pSecDesc, (size_t)cbBuffer);
				}
			}

			//Grab our header for writes
			//header is right at the start of the buffer
			MAPIFreeBuffer(m_lpHeader);
			m_lpHeader = NULL;

			m_cbHeader = CbSecurityDescriptorHeader(lpSDBuffer);

			//make sure we don't try to copy more than we really got
			if (m_cbHeader <= cbSBBuffer)
			{
				EC_H(MAPIAllocateBuffer(
					m_cbHeader,
					(LPVOID*) &m_lpHeader));

				if (m_lpHeader)
				{
					memcpy(m_lpHeader,lpSDBuffer,m_cbHeader);

				}
			}
			//Dump our SD
			CString szDACL;
			CString szInfo;
			EC_H(SDToString(lpSDBuffer,m_acetype,&szDACL,&szInfo));

			DebugPrint(DBGGeneric,_T("sdInfo: %s\nszDACL: %s\n"),szInfo,szDACL);
		}
	}
	MAPIFreeBuffer(lpSDBuffer);

	if (!*ppSecurityDescriptor) return MAPI_E_NOT_FOUND;
	return hRes;
}

//This is very dangerous code and should only be executed under very controlled circumstances
//The code, as written, does nothing to ensure the DACL is ordered correctly, so it will probably cause problems
//on the server once written
//For this reason, the property sheet is read-only unless a reg key is set.
STDMETHODIMP CMySecInfo::SetSecurity(THIS_ SECURITY_INFORMATION /*SecurityInformation*/,
									PSECURITY_DESCRIPTOR pSecurityDescriptor )
{
	DebugPrint(DBGGeneric,_T("CMySecInfo::SetSecurity\n"));
	HRESULT		hRes = S_OK;
	SPropValue	Blob = {0};
	ULONG		cbBlob = 0;
	LPBYTE		lpBlob = NULL;
	DWORD		dwSDLength = 0;

	if (!m_lpHeader || !pSecurityDescriptor || !m_lpMAPIProp) return MAPI_E_INVALID_PARAMETER;
	if (!IsValidSecurityDescriptor(pSecurityDescriptor)) return MAPI_E_INVALID_PARAMETER;

	dwSDLength = ::GetSecurityDescriptorLength(pSecurityDescriptor);
	cbBlob = m_cbHeader + dwSDLength;
	if (cbBlob < m_cbHeader || cbBlob < dwSDLength) return MAPI_E_INVALID_PARAMETER;

	EC_H(MAPIAllocateBuffer(
		cbBlob,
		(LPVOID*) &lpBlob));

	if (lpBlob)
	{
		//The format is a security descriptor preceeded by a header.
		memcpy(lpBlob,m_lpHeader,m_cbHeader);
		EC_B(MakeSelfRelativeSD(pSecurityDescriptor, lpBlob+m_cbHeader, &dwSDLength));

		Blob.ulPropTag = m_ulPropTag;
		Blob.dwAlignPad = NULL;
		Blob.Value.bin.cb = cbBlob;
		Blob.Value.bin.lpb = lpBlob;

		EC_H(HrSetOneProp(
			m_lpMAPIProp,
			&Blob));

		MAPIFreeBuffer(lpBlob);
	}
	return hRes;
}
STDMETHODIMP CMySecInfo::GetAccessRights(THIS_ const GUID* /*pguidObjectType*/,
										DWORD /*dwFlags*/,
										PSI_ACCESS *ppAccess,
										ULONG *pcAccesses,
										ULONG *piDefaultAccess )
{
	DebugPrint(DBGGeneric,_T("CMySecInfo::GetAccessRights\n"));

	switch (m_acetype)
	{
	case acetypeContainer:
		*ppAccess = siExchangeAccessesFolder;
		*pcAccesses = sizeof(siExchangeAccessesFolder)/sizeof(SI_ACCESS);
		break;
	case acetypeMessage:
		*ppAccess = siExchangeAccessesMessage;
		*pcAccesses = sizeof(siExchangeAccessesMessage)/sizeof(SI_ACCESS);
		break;
	case acetypeFreeBusy:
		*ppAccess = siExchangeAccessesFreeBusy;
		*pcAccesses = sizeof(siExchangeAccessesFreeBusy)/sizeof(SI_ACCESS);
		break;
	};

	*piDefaultAccess = 0;
	return S_OK;
};

STDMETHODIMP CMySecInfo::MapGeneric(THIS_ const GUID* /*pguidObjectType*/,
								   UCHAR * /*pAceFlags*/,
								   ACCESS_MASK *pMask)
{
	DebugPrint(DBGGeneric,_T("CMySecInfo::MapGeneric\n"));

	switch (m_acetype)
	{
	case acetypeContainer:
		MapGenericMask(pMask, &gmFolders);
		break;
	case acetypeMessage:
		MapGenericMask(pMask, &gmMessages);
		break;
	case acetypeFreeBusy:
		// No generic for freebusy
		break;
	};
	return S_OK;
};

STDMETHODIMP CMySecInfo::GetInheritTypes(THIS_ PSI_INHERIT_TYPE* /*ppInheritTypes*/,
										ULONG* /*pcInheritTypes*/)
{
	DebugPrint(DBGGeneric,_T("CMySecInfo::GetInheritTypes\n"));
	return E_NOTIMPL;
};

STDMETHODIMP CMySecInfo::PropertySheetPageCallback(THIS_ HWND /*hwnd*/, UINT uMsg, SI_PAGE_TYPE uPage )
{
	DebugPrint(DBGGeneric,_T("CMySecInfo::PropertySheetPageCallback, uMsg = 0x%X, uPage = 0x%X\n"),uMsg,uPage);
	return S_OK;
};

STDMETHODIMP_(BOOL) CMySecInfo::IsDaclCanonical(THIS_ IN PACL /*pDacl*/)
{
	DebugPrint(DBGGeneric,_T("CMySecInfo::IsDaclCanonical - always returns true.\n"));
	return true;
};

STDMETHODIMP CMySecInfo::LookupSids(THIS_ IN ULONG /*cSids*/, IN PSID* /*rgpSids*/, OUT LPDATAOBJECT* /*ppdo*/)
{
	DebugPrint(DBGGeneric,_T("CMySecInfo::LookupSids\n"));
	return E_NOTIMPL;
};

HRESULT DisplayPropAsSD(LPMAPIPROP lpMAPIProp,ULONG ulPropTag)
{
	HRESULT hRes = S_OK;

	DebugPrint(DBGGeneric,_T("DisplayPropAsSD displaying security descriptor from 0x%08X\n"),ulPropTag);
	if (!lpMAPIProp || !ulPropTag) return MAPI_E_INVALID_PARAMETER;

	LPBYTE lpSDToParse = NULL;

	EC_H(GetBuffer(lpMAPIProp,ulPropTag,NULL,&lpSDToParse));

	if (lpSDToParse)
	{
		eAceType acetype = acetypeMessage;
		switch (GetMAPIObjectType(lpMAPIProp))
		{
		case (MAPI_STORE):
		case (MAPI_ADDRBOOK):
		case (MAPI_FOLDER):
		case (MAPI_ABCONT):
			acetype = acetypeContainer;
			break;
		}

		if (PR_FREEBUSY_NT_SECURITY_DESCRIPTOR == ulPropTag)
			acetype = acetypeFreeBusy;

		CString szDACL;
		CString szInfo;

		EC_H(SDToString(lpSDToParse, acetype, &szDACL, &szInfo));

		CEditor MySD(
			NULL,
			IDS_SECDESC,
			IDS_SECDESC,
			3,
			CEDITOR_BUTTON_OK);
		MySD.InitSingleLineSz(0,IDS_SECINFO,szInfo,true);
		MySD.InitSingleLine(1,IDS_SECVERSION,NULL,true);
		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagSecurityVersion, SECURITY_DESCRIPTOR_VERSION(lpSDToParse), &szFlags));
		MySD.SetStringf(1,_T("0x%04X = %s"),SECURITY_DESCRIPTOR_VERSION(lpSDToParse),szFlags);// STRING_OK
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
		MySD.InitMultiLine(2,IDS_DESCRIPTOR,szDACL,true);

		WC_H(MySD.DisplayDialog());
		hRes = S_OK;

		DebugPrint(DBGGeneric,_T("%s\n"),(LPCTSTR) szDACL);
	}
	MAPIFreeBuffer(lpSDToParse);
	return hRes;
};

BOOL GetTextualSid(
				   PSID pSid,            // binary SID
				   LPTSTR TextualSid,    // buffer for Textual representation of SID
				   LPDWORD lpdwBufferLen // required/provided TextualSid buffersize
				   )
{
	HRESULT hRes = S_OK;
	PSID_IDENTIFIER_AUTHORITY psia;
	DWORD dwSubAuthorities;
	DWORD dwSidRev=SID_REVISION;
	DWORD dwCounter = 0;
	DWORD dwSidSize = 0;

	// Validate the binary SID.
	if(!IsValidSid(pSid)) return FALSE;

	// Get the identifier authority value from the SID.
	psia = GetSidIdentifierAuthority(pSid);

	// Get the number of subauthorities in the SID.
	dwSubAuthorities = *GetSidSubAuthorityCount(pSid);

	// Compute the buffer length.
	// S-SID_REVISION- + IdentifierAuthority- + subauthorities- + NULL
	dwSidSize=(15 + 12 + (12 * dwSubAuthorities) + 1) * sizeof(TCHAR);

	// Check input buffer length.
	// If too small, indicate the proper size and set last error.
	if (*lpdwBufferLen < dwSidSize)
	{
		*lpdwBufferLen = dwSidSize;
		SetLastError(ERROR_INSUFFICIENT_BUFFER);
		return FALSE;
	}

	// Add 'S' prefix and revision number to the string.
	EC_H(StringCchPrintf(TextualSid, *lpdwBufferLen, _T("S-%lu-"), dwSidRev ));// STRING_OK

	size_t cchTextualSid = 0;
	EC_H(StringCchLength(TextualSid,STRSAFE_MAX_CCH,&cchTextualSid));

	// Add SID identifier authority to the string.
	if ((psia->Value[0] != 0) || (psia->Value[1] != 0))
	{
		EC_H(StringCchPrintf(TextualSid + cchTextualSid,
			*lpdwBufferLen - cchTextualSid,
			_T("0x%02hx%02hx%02hx%02hx%02hx%02hx"),// STRING_OK
			(USHORT)psia->Value[0],
			(USHORT)psia->Value[1],
			(USHORT)psia->Value[2],
			(USHORT)psia->Value[3],
			(USHORT)psia->Value[4],
			(USHORT)psia->Value[5]));
	}
	else
	{
		EC_H(StringCchPrintf(TextualSid + cchTextualSid,
			*lpdwBufferLen - cchTextualSid,
			_T("%lu"),// STRING_OK
			(ULONG)(psia->Value[5]      ) +
			(ULONG)(psia->Value[4] <<  8) +
			(ULONG)(psia->Value[3] << 16) +
			(ULONG)(psia->Value[2] << 24)));
	}


	// Add SID subauthorities to the string.
	for (dwCounter=0 ; dwCounter < dwSubAuthorities ; dwCounter++)
	{
		EC_H(StringCchLength(TextualSid,STRSAFE_MAX_CCH,&cchTextualSid));
		EC_H(StringCchPrintf(TextualSid + cchTextualSid,
			*lpdwBufferLen - cchTextualSid,
			_T("-%lu"),// STRING_OK
			*GetSidSubAuthority(pSid, dwCounter)));
	}

	return TRUE;
}

HRESULT ACEToString(void* pACE, eAceType acetype, CString *AceString)
{
	HRESULT hRes = S_OK;
	CString szTmpString;
	BYTE	AceType = 0;
	BYTE	AceFlags = 0;
	ACCESS_MASK	Mask = 0;
	DWORD	Flags = 0;
	GUID	ObjectType = {0};
	GUID	InheritedObjectType = {0};
	SID*	SidStart = NULL;
	BOOL	bObjectFound = false;

	if (!pACE || !AceString) return MAPI_E_NOT_FOUND;

	/* Check type of ACE */
	switch (((PACE_HEADER)pACE)->AceType)
	{
	case ACCESS_ALLOWED_ACE_TYPE:
		AceType = ((ACCESS_ALLOWED_ACE *) pACE)->Header.AceType;
		AceFlags = ((ACCESS_ALLOWED_ACE *) pACE)->Header.AceFlags;
		Mask = ((ACCESS_ALLOWED_ACE *) pACE)->Mask;
		SidStart = (SID *) &((ACCESS_ALLOWED_ACE *) pACE)->SidStart;
		break;
	case ACCESS_DENIED_ACE_TYPE:
		AceType = ((ACCESS_DENIED_ACE *) pACE)->Header.AceType;
		AceFlags = ((ACCESS_DENIED_ACE *) pACE)->Header.AceFlags;
		Mask = ((ACCESS_DENIED_ACE *) pACE)->Mask;
		SidStart = (SID *) &((ACCESS_DENIED_ACE *) pACE)->SidStart;
		break;
	case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
		AceType = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->Header.AceType;
		AceFlags = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->Header.AceFlags;
		Mask = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->Mask;
		Flags = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->Flags;
		ObjectType = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->ObjectType;
		InheritedObjectType = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->InheritedObjectType;
		SidStart = (SID *) &((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->SidStart;
		bObjectFound = true;
		break;
	case ACCESS_DENIED_OBJECT_ACE_TYPE:
		AceType = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->Header.AceType;
		AceFlags = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->Header.AceFlags;
		Mask = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->Mask;
		Flags = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->Flags;
		ObjectType = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->ObjectType;
		InheritedObjectType = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->InheritedObjectType;
		SidStart = (SID *) &((ACCESS_DENIED_OBJECT_ACE *) pACE)->SidStart;
		bObjectFound = true;
		break;
	}

	DWORD dwSidName = 0;
	DWORD dwSidDomain = 0;
	SID_NAME_USE SidNameUse;

	WC_B(LookupAccountSid(
		NULL,
		SidStart,
		NULL,
		&dwSidName,
		NULL,
		&dwSidDomain,
		&SidNameUse));
	hRes = S_OK;

	LPTSTR lpSidName = NULL;
	LPTSTR lpSidDomain = NULL;

	if (dwSidName) lpSidName = new TCHAR[dwSidName];
	if (dwSidDomain) lpSidDomain = new TCHAR[dwSidDomain];

	//Only make the call if we got something to get
	if (lpSidName || lpSidDomain)
	{
		WC_B(LookupAccountSid(
			NULL,
			SidStart,
			lpSidName,
			&dwSidName,
			lpSidDomain,
			&dwSidDomain,
			&SidNameUse));
		hRes = S_OK;
	}

	DWORD dwStringSid = 0;
	GetTextualSid(SidStart,NULL,&dwStringSid);//Get a buffer count
	LPTSTR lpStringSid = NULL;
	if (dwStringSid)
	{
		lpStringSid = new TCHAR[dwStringSid];
		if (lpStringSid)
		{
			EC_B(GetTextualSid(SidStart,lpStringSid,&dwStringSid));
		}
	}

	LPTSTR szAceType = NULL;
	LPTSTR szAceFlags = NULL;
	LPTSTR szAceMask = NULL;
	EC_H(InterpretFlags(flagACEType, AceType, &szAceType));
	EC_H(InterpretFlags(flagACEFlag, AceFlags, &szAceFlags));
	switch(acetype)
	{
	case acetypeContainer:
		EC_H(InterpretFlags(flagACEMaskContainer, Mask, &szAceMask));
		break;
	case acetypeMessage:
		EC_H(InterpretFlags(flagACEMaskNonContainer, Mask, &szAceMask));
		break;
	case acetypeFreeBusy:
		EC_H(InterpretFlags(flagACEMaskFreeBusy, Mask, &szAceMask));
		break;
	};
	CString szDomain;
	CString szName;
	CString szSID;

	if (lpSidDomain) szDomain = lpSidDomain;
	else szDomain.LoadString(IDS_NODOMAIN);
	if (lpSidName) szName = lpSidName;
	else szName.LoadString(IDS_NONAME);
	if (lpStringSid) szSID = lpStringSid;
	else szSID.LoadString(IDS_NOSID);

	szTmpString.FormatMessage(
		IDS_SIDACCOUNT,
		szDomain,
		szName,
		szSID,
		AceType, szAceType,
		AceFlags,szAceFlags,
		Mask, szAceMask);
	MAPIFreeBuffer(szAceMask);
	MAPIFreeBuffer(szAceFlags);
	MAPIFreeBuffer(szAceType);
	szAceType = NULL;
	szAceFlags = NULL;
	szAceMask = NULL;
	delete[] lpSidDomain;
	delete[] lpSidName;
	delete[] lpStringSid;

	*AceString += szTmpString;

	if (bObjectFound)
	{
		szTmpString.FormatMessage(IDS_SIDOBJECTYPE);
		*AceString += szTmpString;
		LPTSTR szObjectGuid = GUIDToStringAndName(&ObjectType);
		if (szObjectGuid)
		{
			*AceString += szObjectGuid;
			delete[] szObjectGuid;
		}
		szTmpString.FormatMessage(IDS_SIDINHERITEDOBJECTYPE);
		*AceString += szTmpString;
		LPTSTR szInheritedGuid = GUIDToStringAndName(&InheritedObjectType);
		if (szInheritedGuid)
		{
			*AceString += szInheritedGuid;
			delete[] szInheritedGuid;
		}
		szTmpString.FormatMessage(IDS_SIDFLAGS,Flags);
		*AceString += szTmpString;
	}
	return hRes;
}

HRESULT SDToString(LPBYTE lpBuf, eAceType acetype, CString *SDString, CString *sdInfo)
{
	HRESULT hRes = S_OK;
	BOOL bValidDACL = false;
	PACL pACL = NULL;
	BOOL bDACLDefaulted = false;
	PSECURITY_DESCRIPTOR pSecurityDescriptor = NULL;

	if (!lpBuf || !SDString) return MAPI_E_NOT_FOUND;

	pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(lpBuf);

	if (!IsValidSecurityDescriptor(pSecurityDescriptor))
	{
		SDString->FormatMessage(IDS_INVALIDSD);
		return S_OK;
	}

	if (sdInfo)
	{
		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagSecurityInfo, SECURITY_INFORMATION_OF(lpBuf), &szFlags));
		*sdInfo = szFlags;
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
	}

	EC_B(GetSecurityDescriptorDacl(
		pSecurityDescriptor,
		&bValidDACL,
		&pACL,
		&bDACLDefaulted));
	if (bValidDACL && pACL)
	{
		ACL_SIZE_INFORMATION ACLSizeInfo = {0};
		EC_B(GetAclInformation(
			pACL,
			&ACLSizeInfo,
			sizeof(ACLSizeInfo),
			AclSizeInformation));

		for(DWORD i = 0 ; i < ACLSizeInfo.AceCount ; i++)
		{
			CString szTmpString;
			void*	pACE = NULL;

			EC_B(GetAce(
				pACL,
				i,
				&pACE));

			ACEToString(pACE,acetype,&szTmpString);

			*SDString += szTmpString;
		}
	}
	return hRes;
}
