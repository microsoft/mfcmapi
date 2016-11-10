#pragma once
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
		_Deref_out_opt_ LPVOID* ppvObj) override;
	STDMETHODIMP_(ULONG) AddRef() override;
	STDMETHODIMP_(ULONG) Release() override;

	STDMETHOD(GetObjectInformation) (PSI_OBJECT_INFO pObjectInfo) override;
	STDMETHOD(GetSecurity) (SECURITY_INFORMATION RequestedInformation,
		PSECURITY_DESCRIPTOR* ppSecurityDescriptor,
		BOOL fDefault) override;
	STDMETHOD(SetSecurity) (SECURITY_INFORMATION SecurityInformation,
		PSECURITY_DESCRIPTOR pSecurityDescriptor) override;
	STDMETHOD(GetAccessRights) (const GUID* pguidObjectType,
		DWORD dwFlags,
		PSI_ACCESS* ppAccess,
		ULONG* pcAccesses,
		ULONG* piDefaultAccess) override;
	STDMETHOD(MapGeneric) (const GUID* pguidObjectType,
		UCHAR* pAceFlags,
		ACCESS_MASK *pMask) override;
	STDMETHOD(GetInheritTypes) (PSI_INHERIT_TYPE* ppInheritTypes,
		ULONG* pcInheritTypes) override;
	STDMETHOD(PropertySheetPageCallback)(HWND hwnd, UINT uMsg, SI_PAGE_TYPE uPage) override;
	STDMETHOD_(BOOL, IsDaclCanonical) (PACL pDacl) override;
	STDMETHOD(LookupSids) (ULONG cSids, PSID* rgpSids, LPDATAOBJECT* ppdo) override;

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
_Check_return_ HRESULT SDToString(_In_count_(cbBuf) LPBYTE lpBuf, size_t cbBuf, eAceType acetype, _In_ wstring& SDString, _In_ wstring& sdInfo);
