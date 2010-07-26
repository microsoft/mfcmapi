#pragma once
// MySecInfo.h : header file

#include <Aclui.h>

enum eAceType {
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

	_Check_return_ STDMETHODIMP QueryInterface (REFIID riid,
		_Deref_out_opt_ LPVOID* ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	_Check_return_ STDMETHOD(GetObjectInformation) (_In_ PSI_OBJECT_INFO pObjectInfo);
	_Check_return_ STDMETHOD(GetSecurity) (SECURITY_INFORMATION RequestedInformation,
		_Deref_out_opt_ PSECURITY_DESCRIPTOR* ppSecurityDescriptor,
		BOOL fDefault);
	_Check_return_ STDMETHOD(SetSecurity) (SECURITY_INFORMATION SecurityInformation,
		_Out_ PSECURITY_DESCRIPTOR pSecurityDescriptor );
	_Check_return_ STDMETHOD(GetAccessRights) (_In_ const GUID* pguidObjectType,
		DWORD dwFlags,
		_Out_ PSI_ACCESS* ppAccess,
		_Out_ ULONG* pcAccesses,
		_Out_ ULONG* piDefaultAccess );
	_Check_return_ STDMETHOD(MapGeneric) (_In_ const GUID* pguidObjectType,
		_In_ UCHAR* pAceFlags,
		_Out_ ACCESS_MASK *pMask);
	_Check_return_ STDMETHOD(GetInheritTypes) (_Out_ PSI_INHERIT_TYPE* ppInheritTypes,
		_Out_ ULONG* pcInheritTypes);
	_Check_return_ STDMETHOD(PropertySheetPageCallback)(_In_ HWND hwnd, UINT uMsg, SI_PAGE_TYPE uPage );
	_Check_return_ STDMETHOD_(BOOL,IsDaclCanonical) (_In_ PACL pDacl);
	_Check_return_ STDMETHOD(LookupSids) (ULONG cSids, _In_count_(cSids) PSID* rgpSids, _Deref_out_ LPDATAOBJECT* ppdo);

private:
	LONG		m_cRef;
	LPMAPIPROP	m_lpMAPIProp;
	ULONG		m_ulPropTag;
	LPBYTE		m_lpHeader;
	ULONG		m_cbHeader;
	eAceType	m_acetype;
	WCHAR		m_wszObject[64];
};

_Check_return_ HRESULT SDToString(_In_ LPBYTE lpBuf, eAceType acetype, _In_ CString *SDString, _In_ CString *sdInfo);
