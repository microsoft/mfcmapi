#pragma once
// MySecInfo.h : header file
//
#include <Aclui.h>

enum eAceType {
	acetypeContainer,
	acetypeMessage,
	acetypeFreeBusy
};

class CMySecInfo : public ISecurityInformation, ISecurityInformation2
{
private:
	LONG		m_cRef;
	LPMAPIPROP	m_lpMAPIProp;
	ULONG		m_ulPropTag;
	LPBYTE		m_lpHeader;
	ULONG		m_cbHeader;
	eAceType	m_acetype;
	WCHAR		m_wszObject[64];

public:
	CMySecInfo(LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag);
	~CMySecInfo();

	STDMETHODIMP QueryInterface (REFIID riid,
		LPVOID * ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

    STDMETHOD(GetObjectInformation) (THIS_ PSI_OBJECT_INFO pObjectInfo );
	STDMETHOD(GetSecurity) (THIS_ SECURITY_INFORMATION /*RequestedInformation*/,
                            PSECURITY_DESCRIPTOR *ppSecurityDescriptor,
                            BOOL /*fDefault*/);
    STDMETHOD(SetSecurity) (THIS_ SECURITY_INFORMATION /*SecurityInformation*/,
                            PSECURITY_DESCRIPTOR pSecurityDescriptor );
    STDMETHOD(GetAccessRights) (THIS_ const GUID* /*pguidObjectType*/,
                                DWORD /*dwFlags*/,
                                PSI_ACCESS *ppAccess,
                                ULONG *pcAccesses,
                                ULONG *piDefaultAccess );
    STDMETHOD(MapGeneric) (THIS_ const GUID* /*pguidObjectType*/,
                           UCHAR * /*pAceFlags*/,
                           ACCESS_MASK *pMask);
    STDMETHOD(GetInheritTypes) (THIS_ PSI_INHERIT_TYPE* /*ppInheritTypes*/,
                                ULONG* /*pcInheritTypes*/);
    STDMETHOD(PropertySheetPageCallback)(THIS_ HWND /*hwnd*/, UINT uMsg, SI_PAGE_TYPE uPage );
    STDMETHOD_(BOOL,IsDaclCanonical) (THIS_ IN PACL /*pDacl*/);
    STDMETHOD(LookupSids) (THIS_ IN ULONG /*cSids*/, IN PSID* /*rgpSids*/, OUT LPDATAOBJECT* /*ppdo*/);
};

HRESULT SDToString(LPBYTE lpBuf, eAceType acetype, CString *SDString, CString *sdInfo);
