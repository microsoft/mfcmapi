#include <StdAfx.h>
#include <UI/MySecInfo.h>
#include <Interpret/Sid.h>
#include <MAPI/MAPIFunctions.h>
#include <core/mapi/extraPropTags.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace mapi
{
	namespace mapiui
	{
		static std::wstring CLASS = L"CMySecInfo";

		// The following array defines the permission names for Exchange objects.
		SI_ACCESS siExchangeAccessesFolder[] = {
			{&GUID_NULL, frightsReadAny, MAKEINTRESOURCEW(IDS_ACCESSREADANY), SI_ACCESS_GENERAL},
			{&GUID_NULL, rightsReadOnly, MAKEINTRESOURCEW(IDS_ACCESSREADONLY), SI_ACCESS_GENERAL},
			{&GUID_NULL, frightsCreate, MAKEINTRESOURCEW(IDS_ACCESSCREATE), SI_ACCESS_GENERAL},
			{&GUID_NULL, frightsEditOwned, MAKEINTRESOURCEW(IDS_ACCESSEDITOWN), SI_ACCESS_GENERAL},
			{&GUID_NULL, frightsDeleteOwned, MAKEINTRESOURCEW(IDS_ACCESSDELETEOWN), SI_ACCESS_CONTAINER},
			{&GUID_NULL, frightsEditAny, MAKEINTRESOURCEW(IDS_ACCESSEDITANY), SI_ACCESS_GENERAL},
			{&GUID_NULL, rightsReadWrite, MAKEINTRESOURCEW(IDS_ACCESSREADWRITE), SI_ACCESS_GENERAL},
			{&GUID_NULL, frightsDeleteAny, MAKEINTRESOURCEW(IDS_ACCESSDELETEANY), SI_ACCESS_GENERAL},
			{&GUID_NULL, frightsCreateSubfolder, MAKEINTRESOURCEW(IDS_ACCESSCREATESUBFOLDER), SI_ACCESS_GENERAL},
			{&GUID_NULL, frightsOwner, MAKEINTRESOURCEW(IDS_ACCESSOWNER), SI_ACCESS_GENERAL},
			{&GUID_NULL, frightsVisible, MAKEINTRESOURCEW(IDS_ACCESSVISIBLE), SI_ACCESS_GENERAL},
			{&GUID_NULL, frightsContact, MAKEINTRESOURCEW(IDS_ACCESSCONTACT), SI_ACCESS_GENERAL}};

		SI_ACCESS siExchangeAccessesMessage[] = {
			{&GUID_NULL, fsdrightDelete, MAKEINTRESOURCEW(IDS_ACCESSDELETE), SI_ACCESS_GENERAL},
			{&GUID_NULL, fsdrightReadProperty, MAKEINTRESOURCEW(IDS_ACCESSREADPROPERTY), SI_ACCESS_GENERAL},
			{&GUID_NULL, fsdrightWriteProperty, MAKEINTRESOURCEW(IDS_ACCESSWRITEPROPERTY), SI_ACCESS_GENERAL},
			// { &GUID_NULL, fsdrightCreateMessage, MAKEINTRESOURCEW(IDS_ACCESSCREATEMESSAGE), SI_ACCESS_GENERAL },
			// { &GUID_NULL, fsdrightSaveMessage, MAKEINTRESOURCEW(IDS_ACCESSSAVEMESSAGE), SI_ACCESS_GENERAL },
			// { &GUID_NULL, fsdrightOpenMessage, MAKEINTRESOURCEW(IDS_ACCESSOPENMESSAGE), SI_ACCESS_GENERAL },
			{&GUID_NULL, fsdrightWriteSD, MAKEINTRESOURCEW(IDS_ACCESSWRITESD), SI_ACCESS_GENERAL},
			{&GUID_NULL, fsdrightWriteOwner, MAKEINTRESOURCEW(IDS_ACCESSWRITEOWNER), SI_ACCESS_GENERAL},
			{&GUID_NULL, fsdrightReadControl, MAKEINTRESOURCEW(IDS_ACCESSREADCONTROL), SI_ACCESS_GENERAL},
		};

		SI_ACCESS siExchangeAccessesFreeBusy[] = {
			{&GUID_NULL, fsdrightFreeBusySimple, MAKEINTRESOURCEW(IDS_ACCESSSIMPLEFREEBUSY), SI_ACCESS_GENERAL},
			{&GUID_NULL, fsdrightFreeBusyDetailed, MAKEINTRESOURCEW(IDS_ACCESSDETAILEDFREEBUSY), SI_ACCESS_GENERAL},
		};

		GENERIC_MAPPING gmFolders = {frightsReadAny, frightsEditOwned | frightsDeleteOwned, rightsNone, rightsAll};

		GENERIC_MAPPING gmMessages = {msgrightsGenericRead,
									  msgrightsGenericWrite,
									  msgrightsGenericExecute,
									  msgrightsGenericAll};

		CMySecInfo::CMySecInfo(_In_ const LPMAPIPROP lpMAPIProp, const ULONG ulPropTag)
		{
			TRACE_CONSTRUCTOR(CLASS);
			m_cRef = 1;
			m_ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_BINARY); // An SD must be in a binary prop
			m_lpMAPIProp = lpMAPIProp;
			m_acetype = sid::acetypeMessage;
			m_cbHeader = 0;
			m_lpHeader = nullptr;

			if (m_lpMAPIProp)
			{
				m_lpMAPIProp->AddRef();
				const auto m_ulObjType = GetMAPIObjectType(m_lpMAPIProp);
				switch (m_ulObjType)
				{
				case MAPI_STORE:
				case MAPI_ADDRBOOK:
				case MAPI_FOLDER:
				case MAPI_ABCONT:
					m_acetype = sid::acetypeContainer;
					break;
				}
			}

			if (PR_FREEBUSY_NT_SECURITY_DESCRIPTOR == m_ulPropTag) m_acetype = sid::acetypeFreeBusy;

			m_wszObject = strings::loadstring(IDS_OBJECT);
		}

		CMySecInfo::~CMySecInfo()
		{
			TRACE_DESTRUCTOR(CLASS);
			MAPIFreeBuffer(m_lpHeader);
			if (m_lpMAPIProp) m_lpMAPIProp->Release();
		}

		STDMETHODIMP CMySecInfo::QueryInterface(_In_ REFIID riid, _Deref_out_opt_ LPVOID* ppvObj)
		{
			*ppvObj = nullptr;
			if (riid == IID_ISecurityInformation || riid == IID_IUnknown)
			{
				*ppvObj = static_cast<ISecurityInformation*>(this);
				AddRef();
				return S_OK;
			}

			if (riid == IID_ISecurityInformation2)
			{
				*ppvObj = static_cast<ISecurityInformation2*>(this);
				AddRef();
				return S_OK;
			}

			return E_NOINTERFACE;
		}

		STDMETHODIMP_(ULONG) CMySecInfo::AddRef()
		{
			const auto lCount = InterlockedIncrement(&m_cRef);
			TRACE_ADDREF(CLASS, lCount);
			return lCount;
		}

		STDMETHODIMP_(ULONG) CMySecInfo::Release()
		{
			const auto lCount = InterlockedDecrement(&m_cRef);
			TRACE_RELEASE(CLASS, lCount);
			if (!lCount) delete this;
			return lCount;
		}

		STDMETHODIMP CMySecInfo::GetObjectInformation(const PSI_OBJECT_INFO pObjectInfo)
		{
			output::DebugPrint(DBGGeneric, L"CMySecInfo::GetObjectInformation\n");

			HKEY hRootKey = nullptr;
			WC_W32_S(RegOpenKeyExW(HKEY_CURRENT_USER, registry::RKEY_ROOT, NULL, KEY_READ, &hRootKey));

			auto bAllowEdits = false;
			if (hRootKey)
			{
				bAllowEdits = !!registry::ReadDWORDFromRegistry(
					hRootKey,
					L"AllowUnsupportedSecurityEdits", // STRING_OK
					DWORD(bAllowEdits));
				EC_W32_S(RegCloseKey(hRootKey));
			}

			if (bAllowEdits)
			{
				pObjectInfo->dwFlags = SI_EDIT_PERMS | SI_EDIT_OWNER | SI_ADVANCED |
									   (sid::acetypeContainer == m_acetype ? SI_CONTAINER : 0);
			}
			else
			{
				pObjectInfo->dwFlags =
					SI_READONLY | SI_ADVANCED | (sid::acetypeContainer == m_acetype ? SI_CONTAINER : 0);
			}

			pObjectInfo->pszObjectName = LPWSTR(m_wszObject.c_str()); // Object being edited
			pObjectInfo->pszServerName = nullptr; // specify DC for lookups
			return S_OK;
		}

		STDMETHODIMP CMySecInfo::GetSecurity(
			SECURITY_INFORMATION /*RequestedInformation*/,
			PSECURITY_DESCRIPTOR* ppSecurityDescriptor,
			BOOL /*fDefault*/)
		{
			output::DebugPrint(DBGGeneric, L"CMySecInfo::GetSecurity\n");

			*ppSecurityDescriptor = nullptr;

			const auto lpsProp = GetLargeBinaryProp(m_lpMAPIProp, m_ulPropTag);
			if (lpsProp && PROP_TYPE(lpsProp->ulPropTag) == PT_BINARY && lpsProp->Value.bin.lpb)
			{
				const auto lpSDBuffer = lpsProp->Value.bin.lpb;
				const auto cbSBBuffer = lpsProp->Value.bin.cb;
				const auto pSecDesc =
					SECURITY_DESCRIPTOR_OF(lpSDBuffer); // will be a pointer into lpPropArray, do not free!

				if (IsValidSecurityDescriptor(pSecDesc))
				{
					if (FCheckSecurityDescriptorVersion(lpSDBuffer))
					{
						const int cbBuffer = GetSecurityDescriptorLength(pSecDesc);
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
						m_lpHeader = mapi::allocate<LPBYTE>(m_cbHeader);
						if (m_lpHeader)
						{
							memcpy(m_lpHeader, lpSDBuffer, m_cbHeader);
						}
					}

					// Dump our SD
					auto sd = SDToString(std::vector<BYTE>(lpSDBuffer, lpSDBuffer + cbSBBuffer), m_acetype);
					output::DebugPrint(DBGGeneric, L"sdInfo: %ws\nszDACL: %ws\n", sd.info.c_str(), sd.dacl.c_str());
				}
			}

			MAPIFreeBuffer(lpsProp);

			if (!*ppSecurityDescriptor) return MAPI_E_NOT_FOUND;
			return S_OK;
		}

		// This is very dangerous code and should only be executed under very controlled circumstances
		// The code, as written, does nothing to ensure the DACL is ordered correctly, so it will probably cause problems
		// on the server once written
		// For this reason, the property sheet is read-only unless a reg key is set.
		STDMETHODIMP
		CMySecInfo::SetSecurity(SECURITY_INFORMATION /*SecurityInformation*/, PSECURITY_DESCRIPTOR pSecurityDescriptor)
		{
			output::DebugPrint(DBGGeneric, L"CMySecInfo::SetSecurity\n");
			LPBYTE lpBlob = nullptr;

			if (!m_lpHeader || !pSecurityDescriptor || !m_lpMAPIProp) return MAPI_E_INVALID_PARAMETER;
			if (!IsValidSecurityDescriptor(pSecurityDescriptor)) return MAPI_E_INVALID_PARAMETER;

			auto dwSDLength = GetSecurityDescriptorLength(pSecurityDescriptor);
			const auto cbBlob = m_cbHeader + dwSDLength;
			if (cbBlob < m_cbHeader || cbBlob < dwSDLength) return MAPI_E_INVALID_PARAMETER;

			auto hRes = S_OK;
			lpBlob = mapi::allocate<LPBYTE>(cbBlob);
			if (lpBlob)
			{
				// The format is a security descriptor preceeded by a header.
				memcpy(lpBlob, m_lpHeader, m_cbHeader);
				hRes = EC_B(MakeSelfRelativeSD(pSecurityDescriptor, lpBlob + m_cbHeader, &dwSDLength));

				if (SUCCEEDED(hRes))
				{
					SPropValue Blob = {};
					Blob.ulPropTag = m_ulPropTag;
					Blob.dwAlignPad = NULL;
					Blob.Value.bin.cb = cbBlob;
					Blob.Value.bin.lpb = lpBlob;

					hRes = EC_MAPI(HrSetOneProp(m_lpMAPIProp, &Blob));
				}

				MAPIFreeBuffer(lpBlob);
			}

			return hRes;
		}

		STDMETHODIMP CMySecInfo::GetAccessRights(
			const GUID* /*pguidObjectType*/,
			DWORD /*dwFlags*/,
			PSI_ACCESS* ppAccess,
			ULONG* pcAccesses,
			ULONG* piDefaultAccess)
		{
			output::DebugPrint(DBGGeneric, L"CMySecInfo::GetAccessRights\n");

			switch (m_acetype)
			{
			case sid::acetypeContainer:
				*ppAccess = siExchangeAccessesFolder;
				*pcAccesses = _countof(siExchangeAccessesFolder);
				break;
			case sid::acetypeMessage:
				*ppAccess = siExchangeAccessesMessage;
				*pcAccesses = _countof(siExchangeAccessesMessage);
				break;
			case sid::acetypeFreeBusy:
				*ppAccess = siExchangeAccessesFreeBusy;
				*pcAccesses = _countof(siExchangeAccessesFreeBusy);
				break;
			};

			*piDefaultAccess = 0;
			return S_OK;
		}

		STDMETHODIMP CMySecInfo::MapGeneric(const GUID* /*pguidObjectType*/, UCHAR* /*pAceFlags*/, ACCESS_MASK* pMask)
		{
			output::DebugPrint(DBGGeneric, L"CMySecInfo::MapGeneric\n");

			switch (m_acetype)
			{
			case sid::acetypeContainer:
				MapGenericMask(pMask, &gmFolders);
				break;
			case sid::acetypeMessage:
				MapGenericMask(pMask, &gmMessages);
				break;
			case sid::acetypeFreeBusy:
				// No generic for freebusy
				break;
			};
			return S_OK;
		}

		STDMETHODIMP CMySecInfo::GetInheritTypes(PSI_INHERIT_TYPE* /*ppInheritTypes*/, ULONG* /*pcInheritTypes*/)
		{
			output::DebugPrint(DBGGeneric, L"CMySecInfo::GetInheritTypes\n");
			return E_NOTIMPL;
		}

		STDMETHODIMP CMySecInfo::PropertySheetPageCallback(HWND /*hwnd*/, const UINT uMsg, const SI_PAGE_TYPE uPage)
		{
			output::DebugPrint(
				DBGGeneric, L"CMySecInfo::PropertySheetPageCallback, uMsg = 0x%X, uPage = 0x%X\n", uMsg, uPage);
			return S_OK;
		}

		STDMETHODIMP_(BOOL) CMySecInfo::IsDaclCanonical(PACL /*pDacl*/)
		{
			output::DebugPrint(DBGGeneric, L"CMySecInfo::IsDaclCanonical - always returns true.\n");
			return true;
		}

		STDMETHODIMP CMySecInfo::LookupSids(ULONG /*cSids*/, PSID* /*rgpSids*/, LPDATAOBJECT* /*ppdo*/)
		{
			output::DebugPrint(DBGGeneric, L"CMySecInfo::LookupSids\n");
			return E_NOTIMPL;
		}
	} // namespace mapiui
} // namespace mapi