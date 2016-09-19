#pragma once
// MySecInfo.h : header file

#include <Aclui.h>

enum eAceType
{
	acetypeContainer,
	acetypeMessage,
	acetypeFreeBusy
};

class CMySecInfo : public ISecurityInformation, ISecurityInformation2
{
public:
	CMySecInfo(_In_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag);
	virtual ~CMySecInfo();

	STDMETHODIMP QueryInterface(_In_ REFIID riid,
		_Deref_out_opt_ LPVOID* ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHOD(GetObjectInformation) (PSI_OBJECT_INFO pObjectInfo);
	STDMETHOD(GetSecurity) (SECURITY_INFORMATION RequestedInformation,
		PSECURITY_DESCRIPTOR* ppSecurityDescriptor,
		BOOL fDefault);
	STDMETHOD(SetSecurity) (SECURITY_INFORMATION SecurityInformation,
		PSECURITY_DESCRIPTOR pSecurityDescriptor);
	STDMETHOD(GetAccessRights) (const GUID* pguidObjectType,
		DWORD dwFlags,
		PSI_ACCESS* ppAccess,
		ULONG* pcAccesses,
		ULONG* piDefaultAccess);
	STDMETHOD(MapGeneric) (const GUID* pguidObjectType,
		UCHAR* pAceFlags,
		ACCESS_MASK *pMask);
	STDMETHOD(GetInheritTypes) (PSI_INHERIT_TYPE* ppInheritTypes,
		ULONG* pcInheritTypes);
	STDMETHOD(PropertySheetPageCallback)(HWND hwnd, UINT uMsg, SI_PAGE_TYPE uPage);
	STDMETHOD_(BOOL, IsDaclCanonical) (PACL pDacl);
	STDMETHOD(LookupSids) (ULONG cSids, PSID* rgpSids, LPDATAOBJECT* ppdo);

private:
	LONG m_cRef;
	LPMAPIPROP m_lpMAPIProp;
	ULONG m_ulPropTag;
	LPBYTE m_lpHeader;
	ULONG m_cbHeader;
	eAceType m_acetype;
	wstring m_wszObject;
};

_Check_return_ wstring GetTextualSid(_In_ PSID pSid);
_Check_return_ HRESULT SDToString(_In_count_(cbBuf) LPBYTE lpBuf, ULONG cbBuf, eAceType acetype, _In_ wstring& SDString, _In_ wstring& sdInfo);