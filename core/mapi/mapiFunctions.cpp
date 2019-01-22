// Collection of useful MAPI functions
#include <core/stdafx.h>
#include <core/mapi/mapiFunctions.h>
#include <core/interpret/guid.h>
#include <core/utility/output.h>
#include <core/utility/error.h>

namespace mapi
{
	// Safely cast across MAPI interfaces. Result is addrefed and must be released.
	template <class T> T safe_cast(IUnknown* src)
	{
		if (!src) return nullptr;

		auto iid = IID();
		// clang-format off
		if (std::is_same_v<T, LPUNKNOWN>) iid = IID_IUnknown;
		else if (std::is_same_v<T, LPMAPIFOLDER>) iid = IID_IMAPIFolder;
		else if (std::is_same_v<T, LPMAPICONTAINER>) iid = IID_IMAPIContainer;
		else if (std::is_same_v<T, LPMAILUSER>) iid = IID_IMailUser;
		else if (std::is_same_v<T, LPABCONT>) iid = IID_IABContainer;
		else if (std::is_same_v<T, LPMESSAGE>) iid = IID_IMessage;
		else if (std::is_same_v<T, LPMDB>) iid = IID_IMsgStore;
		else if (std::is_same_v<T, LPMAPIFORMINFO>) iid = IID_IMAPIFormInfo;
		else if (std::is_same_v<T, LPMAPIPROP>) iid = IID_IMAPIProp;
		else if (std::is_same_v<T, LPMAPIFORM>) iid = IID_IMAPIForm;
		else if (std::is_same_v<T, LPPERSISTMESSAGE>) iid = IID_IPersistMessage;
		else if (std::is_same_v<T, IAttachmentSecurity*>) iid = guid::IID_IAttachmentSecurity;
		else if (std::is_same_v<T, LPSERVICEADMIN2>) iid = IID_IMsgServiceAdmin2;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTORE>) iid = IID_IExchangeManageStore;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTORE3>) iid = IID_IExchangeManageStore3;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTORE4>) iid = IID_IExchangeManageStore4;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTORE5>) iid = guid::IID_IExchangeManageStore5;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTOREEX>) iid = guid::IID_IExchangeManageStoreEx;
		else if (std::is_same_v<T, IProxyStoreObject*>) iid = guid::IID_IProxyStoreObject;
		else if (std::is_same_v<T, LPMAPICLIENTSHUTDOWN>) iid = IID_IMAPIClientShutdown;
		else if (std::is_same_v<T, LPPROFSECT>) iid = IID_IProfSect;
		else if (std::is_same_v<T, LPATTACH>) iid = IID_IAttachment;
		else assert(false);
		// clang-format on

		T ret = nullptr;
		WC_H_S(src->QueryInterface(iid, reinterpret_cast<LPVOID*>(&ret)));
		output::DebugPrint(
			DBGGeneric,
			L"safe_cast: iid =%ws, src = %p, ret = %p\n",
			guid::GUIDToStringAndName(&iid).c_str(),
			src,
			ret);

		return ret;
	}

	template LPUNKNOWN safe_cast<LPUNKNOWN>(IUnknown* src);
	template LPMAPIFOLDER safe_cast<LPMAPIFOLDER>(IUnknown* src);
	template LPMAPICONTAINER safe_cast<LPMAPICONTAINER>(IUnknown* src);
	template LPMAILUSER safe_cast<LPMAILUSER>(IUnknown* src);
	template LPABCONT safe_cast<LPABCONT>(IUnknown* src);
	template LPMESSAGE safe_cast<LPMESSAGE>(IUnknown* src);
	template LPMDB safe_cast<LPMDB>(IUnknown* src);
	template LPMAPIFORMINFO safe_cast<LPMAPIFORMINFO>(IUnknown* src);
	template LPMAPIPROP safe_cast<LPMAPIPROP>(IUnknown* src);
	template LPMAPIFORM safe_cast<LPMAPIFORM>(IUnknown* src);
	template LPPERSISTMESSAGE safe_cast<LPPERSISTMESSAGE>(IUnknown* src);
	template IAttachmentSecurity* safe_cast<IAttachmentSecurity*>(IUnknown* src);
	template LPSERVICEADMIN2 safe_cast<LPSERVICEADMIN2>(IUnknown* src);
	template LPEXCHANGEMANAGESTORE safe_cast<LPEXCHANGEMANAGESTORE>(IUnknown* src);
	template LPEXCHANGEMANAGESTORE3 safe_cast<LPEXCHANGEMANAGESTORE3>(IUnknown* src);
	template LPEXCHANGEMANAGESTORE4 safe_cast<LPEXCHANGEMANAGESTORE4>(IUnknown* src);
	template LPEXCHANGEMANAGESTORE5 safe_cast<LPEXCHANGEMANAGESTORE5>(IUnknown* src);
	template LPEXCHANGEMANAGESTOREEX safe_cast<LPEXCHANGEMANAGESTOREEX>(IUnknown* src);
	template IProxyStoreObject* safe_cast<IProxyStoreObject*>(IUnknown* src);
	template LPMAPICLIENTSHUTDOWN safe_cast<LPMAPICLIENTSHUTDOWN>(IUnknown* src);
	template LPPROFSECT safe_cast<LPPROFSECT>(IUnknown* src);
	template LPATTACH safe_cast<LPATTACH>(IUnknown* src);
} // namespace mapi