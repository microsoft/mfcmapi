#include "stdafx.h"
#include "MySecInfo.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"
#include "String.h"

static wstring CLASS = L"CMySecInfo";

// The following array defines the permission names for Exchange objects.
SI_ACCESS siExchangeAccessesFolder[] =
{
 { &GUID_NULL, frightsReadAny, MAKEINTRESOURCEW(IDS_ACCESSREADANY), SI_ACCESS_GENERAL },
 { &GUID_NULL, rightsReadOnly, MAKEINTRESOURCEW(IDS_ACCESSREADONLY), SI_ACCESS_GENERAL },
 { &GUID_NULL, frightsCreate, MAKEINTRESOURCEW(IDS_ACCESSCREATE), SI_ACCESS_GENERAL },
 { &GUID_NULL, frightsEditOwned, MAKEINTRESOURCEW(IDS_ACCESSEDITOWN), SI_ACCESS_GENERAL },
 { &GUID_NULL, frightsDeleteOwned, MAKEINTRESOURCEW(IDS_ACCESSDELETEOWN), SI_ACCESS_CONTAINER },
 { &GUID_NULL, frightsEditAny, MAKEINTRESOURCEW(IDS_ACCESSEDITANY), SI_ACCESS_GENERAL },
 { &GUID_NULL, rightsReadWrite, MAKEINTRESOURCEW(IDS_ACCESSREADWRITE), SI_ACCESS_GENERAL },
 { &GUID_NULL, frightsDeleteAny, MAKEINTRESOURCEW(IDS_ACCESSDELETEANY), SI_ACCESS_GENERAL },
 { &GUID_NULL, frightsCreateSubfolder, MAKEINTRESOURCEW(IDS_ACCESSCREATESUBFOLDER), SI_ACCESS_GENERAL },
 { &GUID_NULL, frightsOwner, MAKEINTRESOURCEW(IDS_ACCESSOWNER), SI_ACCESS_GENERAL },
 { &GUID_NULL, frightsVisible, MAKEINTRESOURCEW(IDS_ACCESSVISIBLE), SI_ACCESS_GENERAL },
 { &GUID_NULL, frightsContact, MAKEINTRESOURCEW(IDS_ACCESSCONTACT), SI_ACCESS_GENERAL }
};

SI_ACCESS siExchangeAccessesMessage[] =
{
 { &GUID_NULL, fsdrightDelete, MAKEINTRESOURCEW(IDS_ACCESSDELETE), SI_ACCESS_GENERAL },
 { &GUID_NULL, fsdrightReadProperty, MAKEINTRESOURCEW(IDS_ACCESSREADPROPERTY), SI_ACCESS_GENERAL },
 { &GUID_NULL, fsdrightWriteProperty, MAKEINTRESOURCEW(IDS_ACCESSWRITEPROPERTY), SI_ACCESS_GENERAL },
 // { &GUID_NULL, fsdrightCreateMessage, MAKEINTRESOURCEW(IDS_ACCESSCREATEMESSAGE), SI_ACCESS_GENERAL },
 // { &GUID_NULL, fsdrightSaveMessage, MAKEINTRESOURCEW(IDS_ACCESSSAVEMESSAGE), SI_ACCESS_GENERAL },
 // { &GUID_NULL, fsdrightOpenMessage, MAKEINTRESOURCEW(IDS_ACCESSOPENMESSAGE), SI_ACCESS_GENERAL },
 { &GUID_NULL, fsdrightWriteSD, MAKEINTRESOURCEW(IDS_ACCESSWRITESD), SI_ACCESS_GENERAL },
 { &GUID_NULL, fsdrightWriteOwner, MAKEINTRESOURCEW(IDS_ACCESSWRITEOWNER), SI_ACCESS_GENERAL },
 { &GUID_NULL, fsdrightReadControl, MAKEINTRESOURCEW(IDS_ACCESSREADCONTROL), SI_ACCESS_GENERAL },
};

SI_ACCESS siExchangeAccessesFreeBusy[] =
{
 { &GUID_NULL, fsdrightFreeBusySimple, MAKEINTRESOURCEW(IDS_ACCESSSIMPLEFREEBUSY), SI_ACCESS_GENERAL },
 { &GUID_NULL, fsdrightFreeBusyDetailed, MAKEINTRESOURCEW(IDS_ACCESSDETAILEDFREEBUSY), SI_ACCESS_GENERAL },
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

CMySecInfo::CMySecInfo(_In_ LPMAPIPROP lpMAPIProp,
	ULONG ulPropTag)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_cRef = 1;
	m_ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_BINARY); // An SD must be in a binary prop
	m_lpMAPIProp = lpMAPIProp;
	m_acetype = acetypeMessage;
	m_cbHeader = 0;
	m_lpHeader = nullptr;

	if (m_lpMAPIProp)
	{
		m_lpMAPIProp->AddRef();
		auto m_ulObjType = GetMAPIObjectType(m_lpMAPIProp);
		switch (m_ulObjType)
		{
		case MAPI_STORE:
		case MAPI_ADDRBOOK:
		case MAPI_FOLDER:
		case MAPI_ABCONT:
			m_acetype = acetypeContainer;
			break;
		}
	}

	if (PR_FREEBUSY_NT_SECURITY_DESCRIPTOR == m_ulPropTag)
		m_acetype = acetypeFreeBusy;

	m_wszObject = loadstring(IDS_OBJECT);
}

CMySecInfo::~CMySecInfo()
{
	TRACE_DESTRUCTOR(CLASS);
	MAPIFreeBuffer(m_lpHeader);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
}

STDMETHODIMP CMySecInfo::QueryInterface(_In_ REFIID riid,
	_Deref_out_opt_ LPVOID* ppvObj)
{
	*ppvObj = nullptr;
	if (riid == IID_ISecurityInformation ||
		riid == IID_IUnknown)
	{
		*ppvObj = static_cast<ISecurityInformation *>(this);
		AddRef();
		return S_OK;
	}

	if (riid == IID_ISecurityInformation2)
	{
		*ppvObj = static_cast<ISecurityInformation2 *>(this);
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CMySecInfo::AddRef()
{
	auto lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS, lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CMySecInfo::Release()
{
	auto lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS, lCount);
	if (!lCount) delete this;
	return lCount;
}

STDMETHODIMP CMySecInfo::GetObjectInformation(PSI_OBJECT_INFO pObjectInfo)
{
	DebugPrint(DBGGeneric, L"CMySecInfo::GetObjectInformation\n");
	auto hRes = S_OK;
	auto bAllowEdits = false;

	HKEY hRootKey = nullptr;

	WC_W32(RegOpenKeyExW(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		NULL,
		KEY_READ,
		&hRootKey));

	if (hRootKey)
	{
		bAllowEdits = !!ReadDWORDFromRegistry(
			hRootKey,
			L"AllowUnsupportedSecurityEdits", // STRING_OK
			DWORD(bAllowEdits));
		EC_W32(RegCloseKey(hRootKey));
	}

	if (bAllowEdits)
	{
		pObjectInfo->dwFlags = SI_EDIT_PERMS | SI_EDIT_OWNER | SI_ADVANCED | (acetypeContainer == m_acetype ? SI_CONTAINER : 0);
	}
	else
	{
		pObjectInfo->dwFlags = SI_READONLY | SI_ADVANCED | (acetypeContainer == m_acetype ? SI_CONTAINER : 0);
	}
	pObjectInfo->pszObjectName = LPWSTR(m_wszObject.c_str()); // Object being edited
	pObjectInfo->pszServerName = nullptr; // specify DC for lookups
	return S_OK;
}

STDMETHODIMP CMySecInfo::GetSecurity(SECURITY_INFORMATION /*RequestedInformation*/,
	PSECURITY_DESCRIPTOR *ppSecurityDescriptor,
	BOOL /*fDefault*/)
{
	DebugPrint(DBGGeneric, L"CMySecInfo::GetSecurity\n");
	auto hRes = S_OK;
	LPSPropValue lpsProp = nullptr;

	*ppSecurityDescriptor = nullptr;

	EC_H(GetLargeBinaryProp(m_lpMAPIProp, m_ulPropTag, &lpsProp));

	if (lpsProp && PROP_TYPE(lpsProp->ulPropTag) == PT_BINARY && lpsProp->Value.bin.lpb)
	{
		auto lpSDBuffer = lpsProp->Value.bin.lpb;
		auto cbSBBuffer = lpsProp->Value.bin.cb;
		auto pSecDesc = SECURITY_DESCRIPTOR_OF(lpSDBuffer); // will be a pointer into lpPropArray, do not free!

		if (IsValidSecurityDescriptor(pSecDesc))
		{
			if (FCheckSecurityDescriptorVersion(lpSDBuffer))
			{
				int cbBuffer = GetSecurityDescriptorLength(pSecDesc);
				*ppSecurityDescriptor = LocalAlloc(LMEM_FIXED, cbBuffer);

				if (*ppSecurityDescriptor)
				{
					memcpy(*ppSecurityDescriptor, pSecDesc, static_cast<size_t>(cbBuffer));
				}
			}

			// Grab our header for writes
			// header is right at the start of the buffer
			MAPIFreeBuffer(m_lpHeader);
			m_lpHeader = nullptr;
			m_cbHeader = CbSecurityDescriptorHeader(lpSDBuffer);

			// make sure we don't try to copy more than we really got
			if (m_cbHeader <= cbSBBuffer)
			{
				EC_H(MAPIAllocateBuffer(
					m_cbHeader,
					reinterpret_cast<LPVOID*>(&m_lpHeader)));

				if (m_lpHeader)
				{
					memcpy(m_lpHeader, lpSDBuffer, m_cbHeader);

				}
			}

			// Dump our SD
			wstring szDACL;
			wstring szInfo;
			EC_H(SDToString(lpSDBuffer, cbSBBuffer, m_acetype, szDACL, szInfo));

			DebugPrint(DBGGeneric, L"sdInfo: %ws\nszDACL: %ws\n", szInfo.c_str(), szDACL.c_str());
		}
	}
	MAPIFreeBuffer(lpsProp);

	if (!*ppSecurityDescriptor) return MAPI_E_NOT_FOUND;
	return hRes;
}

// This is very dangerous code and should only be executed under very controlled circumstances
// The code, as written, does nothing to ensure the DACL is ordered correctly, so it will probably cause problems
// on the server once written
// For this reason, the property sheet is read-only unless a reg key is set.
STDMETHODIMP CMySecInfo::SetSecurity(SECURITY_INFORMATION /*SecurityInformation*/,
	PSECURITY_DESCRIPTOR pSecurityDescriptor)
{
	DebugPrint(DBGGeneric, L"CMySecInfo::SetSecurity\n");
	auto hRes = S_OK;
	SPropValue Blob = { 0 };
	LPBYTE lpBlob = nullptr;

	if (!m_lpHeader || !pSecurityDescriptor || !m_lpMAPIProp) return MAPI_E_INVALID_PARAMETER;
	if (!IsValidSecurityDescriptor(pSecurityDescriptor)) return MAPI_E_INVALID_PARAMETER;

	auto dwSDLength = ::GetSecurityDescriptorLength(pSecurityDescriptor);
	auto cbBlob = m_cbHeader + dwSDLength;
	if (cbBlob < m_cbHeader || cbBlob < dwSDLength) return MAPI_E_INVALID_PARAMETER;

	EC_H(MAPIAllocateBuffer(
		cbBlob,
		reinterpret_cast<LPVOID*>(&lpBlob)));

	if (lpBlob)
	{
		// The format is a security descriptor preceeded by a header.
		memcpy(lpBlob, m_lpHeader, m_cbHeader);
		EC_B(MakeSelfRelativeSD(pSecurityDescriptor, lpBlob + m_cbHeader, &dwSDLength));

		Blob.ulPropTag = m_ulPropTag;
		Blob.dwAlignPad = NULL;
		Blob.Value.bin.cb = cbBlob;
		Blob.Value.bin.lpb = lpBlob;

		EC_MAPI(HrSetOneProp(
			m_lpMAPIProp,
			&Blob));

		MAPIFreeBuffer(lpBlob);
	}
	return hRes;
}

STDMETHODIMP CMySecInfo::GetAccessRights(const GUID* /*pguidObjectType*/,
	DWORD /*dwFlags*/,
	PSI_ACCESS *ppAccess,
	ULONG *pcAccesses,
	ULONG *piDefaultAccess)
{
	DebugPrint(DBGGeneric, L"CMySecInfo::GetAccessRights\n");

	switch (m_acetype)
	{
	case acetypeContainer:
		*ppAccess = siExchangeAccessesFolder;
		*pcAccesses = _countof(siExchangeAccessesFolder);
		break;
	case acetypeMessage:
		*ppAccess = siExchangeAccessesMessage;
		*pcAccesses = _countof(siExchangeAccessesMessage);
		break;
	case acetypeFreeBusy:
		*ppAccess = siExchangeAccessesFreeBusy;
		*pcAccesses = _countof(siExchangeAccessesFreeBusy);
		break;
	};

	*piDefaultAccess = 0;
	return S_OK;
}

STDMETHODIMP CMySecInfo::MapGeneric(const GUID* /*pguidObjectType*/,
	UCHAR* /*pAceFlags*/,
	ACCESS_MASK *pMask)
{
	DebugPrint(DBGGeneric, L"CMySecInfo::MapGeneric\n");

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
}

STDMETHODIMP CMySecInfo::GetInheritTypes(PSI_INHERIT_TYPE* /*ppInheritTypes*/,
	ULONG* /*pcInheritTypes*/)
{
	DebugPrint(DBGGeneric, L"CMySecInfo::GetInheritTypes\n");
	return E_NOTIMPL;
}

STDMETHODIMP CMySecInfo::PropertySheetPageCallback(HWND /*hwnd*/, UINT uMsg, SI_PAGE_TYPE uPage)
{
	DebugPrint(DBGGeneric, L"CMySecInfo::PropertySheetPageCallback, uMsg = 0x%X, uPage = 0x%X\n", uMsg, uPage);
	return S_OK;
}

STDMETHODIMP_(BOOL) CMySecInfo::IsDaclCanonical(PACL /*pDacl*/)
{
	DebugPrint(DBGGeneric, L"CMySecInfo::IsDaclCanonical - always returns true.\n");
	return true;
}

STDMETHODIMP CMySecInfo::LookupSids(ULONG /*cSids*/, PSID* /*rgpSids*/, LPDATAOBJECT* /*ppdo*/)
{
	DebugPrint(DBGGeneric, L"CMySecInfo::LookupSids\n");
	return E_NOTIMPL;
}

_Check_return_ wstring GetTextualSid(_In_ PSID pSid)
{
	// Validate the binary SID.
	if (!IsValidSid(pSid)) return L"";

	// Get the identifier authority value from the SID.
	auto psia = GetSidIdentifierAuthority(pSid);

	// Get the number of subauthorities in the SID.
	auto lpSubAuthoritiesCount = GetSidSubAuthorityCount(pSid);

	// Compute the buffer length.
	// S-SID_REVISION- + IdentifierAuthority- + subauthorities- + NULL
	// Add 'S' prefix and revision number to the string.
	auto TextualSid = format(L"S-%lu-", SID_REVISION); // STRING_OK

	// Add SID identifier authority to the string.
	if (psia->Value[0] != 0 || psia->Value[1] != 0)
	{
		TextualSid += format(
			L"0x%02hx%02hx%02hx%02hx%02hx%02hx", // STRING_OK
			static_cast<USHORT>(psia->Value[0]),
			static_cast<USHORT>(psia->Value[1]),
			static_cast<USHORT>(psia->Value[2]),
			static_cast<USHORT>(psia->Value[3]),
			static_cast<USHORT>(psia->Value[4]),
			static_cast<USHORT>(psia->Value[5]));
	}
	else
	{
		TextualSid += format(
			L"%lu", // STRING_OK
			static_cast<ULONG>(psia->Value[4] << 8) +
			static_cast<ULONG>(psia->Value[5]) +
			static_cast<ULONG>(psia->Value[3] << 16) +
			static_cast<ULONG>(psia->Value[2] << 24));
	}

	// Add SID subauthorities to the string.
	if (lpSubAuthoritiesCount)
	{
		for (DWORD dwCounter = 0; dwCounter < *lpSubAuthoritiesCount; dwCounter++)
		{
			TextualSid += format(
				L"-%lu", // STRING_OK
				*GetSidSubAuthority(pSid, dwCounter));
		}
	}

	return TextualSid;
}

wstring ACEToString(_In_ void* pACE, eAceType acetype)
{
	auto hRes = S_OK;
	wstring AceString;
	ACCESS_MASK Mask = 0;
	DWORD Flags = 0;
	GUID ObjectType = { 0 };
	GUID InheritedObjectType = { 0 };
	SID* SidStart = nullptr;
	auto bObjectFound = false;

	if (!pACE) return L"";

	auto AceType = static_cast<PACE_HEADER>(pACE)->AceType;
	auto AceFlags = static_cast<PACE_HEADER>(pACE)->AceFlags;

	/* Check type of ACE */
	switch (AceType)
	{
	case ACCESS_ALLOWED_ACE_TYPE:
		Mask = static_cast<ACCESS_ALLOWED_ACE *>(pACE)->Mask;
		SidStart = reinterpret_cast<SID *>(&static_cast<ACCESS_ALLOWED_ACE *>(pACE)->SidStart);
		break;
	case ACCESS_DENIED_ACE_TYPE:
		Mask = static_cast<ACCESS_DENIED_ACE *>(pACE)->Mask;
		SidStart = reinterpret_cast<SID *>(&static_cast<ACCESS_DENIED_ACE *>(pACE)->SidStart);
		break;
	case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
		Mask = static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->Mask;
		Flags = static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->Flags;
		ObjectType = static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->ObjectType;
		InheritedObjectType = static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->InheritedObjectType;
		SidStart = reinterpret_cast<SID *>(&static_cast<ACCESS_ALLOWED_OBJECT_ACE *>(pACE)->SidStart);
		bObjectFound = true;
		break;
	case ACCESS_DENIED_OBJECT_ACE_TYPE:
		Mask = static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->Mask;
		Flags = static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->Flags;
		ObjectType = static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->ObjectType;
		InheritedObjectType = static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->InheritedObjectType;
		SidStart = reinterpret_cast<SID *>(&static_cast<ACCESS_DENIED_OBJECT_ACE *>(pACE)->SidStart);
		bObjectFound = true;
		break;
	}

	DWORD dwSidName = 0;
	DWORD dwSidDomain = 0;
	SID_NAME_USE SidNameUse;

	WC_B(LookupAccountSidW(
		NULL,
		SidStart,
		NULL,
		&dwSidName,
		NULL,
		&dwSidDomain,
		&SidNameUse));
	hRes = S_OK;

	LPWSTR lpSidName = nullptr;
	LPWSTR lpSidDomain = nullptr;

#pragma warning(push)
#pragma warning(disable:6211)
	if (dwSidName) lpSidName = new WCHAR[dwSidName];
	if (dwSidDomain) lpSidDomain = new WCHAR[dwSidDomain];
#pragma warning(pop)

	// Only make the call if we got something to get
	if (lpSidName || lpSidDomain)
	{
		WC_B(LookupAccountSidW(
			NULL,
			SidStart,
			lpSidName,
			&dwSidName,
			lpSidDomain,
			&dwSidDomain,
			&SidNameUse));
	}

	auto lpStringSid = GetTextualSid(SidStart);
	auto szAceType = InterpretFlags(flagACEType, AceType);
	auto szAceFlags = InterpretFlags(flagACEFlag, AceFlags);
	wstring szAceMask;

	switch (acetype)
	{
	case acetypeContainer:
		szAceMask = InterpretFlags(flagACEMaskContainer, Mask);
		break;
	case acetypeMessage:
		szAceMask = InterpretFlags(flagACEMaskNonContainer, Mask);
		break;
	case acetypeFreeBusy:
		szAceMask = InterpretFlags(flagACEMaskFreeBusy, Mask);
		break;
	};

	auto szDomain = lpSidDomain ? lpSidDomain : formatmessage(IDS_NODOMAIN);
	auto szName = lpSidName ? lpSidName : formatmessage(IDS_NONAME);
	auto szSID = GetTextualSid(SidStart);
	if (szSID.empty()) szSID = formatmessage(IDS_NOSID);
	delete[] lpSidDomain;
	delete[] lpSidName;

	AceString += formatmessage(
		IDS_SIDACCOUNT,
		szDomain.c_str(),
		szName.c_str(),
		szSID.c_str(),
		AceType, szAceType.c_str(),
		AceFlags, szAceFlags.c_str(),
		Mask, szAceMask.c_str());

	if (bObjectFound)
	{
		AceString += formatmessage(IDS_SIDOBJECTYPE);
		AceString += GUIDToStringAndName(&ObjectType);
		AceString += formatmessage(IDS_SIDINHERITEDOBJECTYPE);
		AceString += GUIDToStringAndName(&InheritedObjectType);
		AceString += formatmessage(IDS_SIDFLAGS, Flags);
	}

	return AceString;
}

_Check_return_ HRESULT SDToString(_In_count_(cbBuf) const BYTE* lpBuf, size_t cbBuf, eAceType acetype, _In_ wstring& SDString, _In_ wstring& sdInfo)
{
	auto hRes = S_OK;
	BOOL bValidDACL = false;
	PACL pACL = nullptr;
	BOOL bDACLDefaulted = false;

	if (!lpBuf) return MAPI_E_NOT_FOUND;

	auto pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(lpBuf);

	if (CbSecurityDescriptorHeader(lpBuf) > cbBuf || !IsValidSecurityDescriptor(pSecurityDescriptor))
	{
		SDString = formatmessage(IDS_INVALIDSD);
		return S_OK;
	}

	sdInfo = InterpretFlags(flagSecurityInfo, SECURITY_INFORMATION_OF(lpBuf));

	EC_B(GetSecurityDescriptorDacl(
		pSecurityDescriptor,
		&bValidDACL,
		&pACL,
		&bDACLDefaulted));
	if (bValidDACL && pACL)
	{
		ACL_SIZE_INFORMATION ACLSizeInfo = { 0 };
		EC_B(GetAclInformation(
			pACL,
			&ACLSizeInfo,
			sizeof ACLSizeInfo,
			AclSizeInformation));

		for (DWORD i = 0; i < ACLSizeInfo.AceCount; i++)
		{
			void* pACE = nullptr;

			EC_B(GetAce(pACL, i, &pACE));

			if (pACE)
			{
				SDString += ACEToString(pACE, acetype);
			}
		}
	}

	return hRes;
}