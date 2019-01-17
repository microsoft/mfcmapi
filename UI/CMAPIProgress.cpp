#include <StdAfx.h>
#include <UI/CMAPIProgress.h>
#include <Enums.h>
#include <UI/Dialogs/BaseDialog.h>
#include <IO/Registry.h>
#include <core/utility/strings.h>
#include <IO/output.h>

namespace ui
{
	static std::wstring CLASS = L"CMAPIProgress";

	_Check_return_ LPMAPIPROGRESS GetMAPIProgress(const std::wstring& lpszContext, _In_ HWND hWnd)
	{
		if (registry::useIMAPIProgress)
		{
			const auto pProgress = new CMAPIProgress(lpszContext, hWnd);

			return static_cast<LPMAPIPROGRESS>(pProgress);
		}

		return nullptr;
	}

	CMAPIProgress::CMAPIProgress(const std::wstring& lpszContext, _In_ HWND hWnd)
	{
		TRACE_CONSTRUCTOR(CLASS);

		m_cRef = 1;
		m_ulMin = 1;
		m_ulMax = 1000;
		m_ulFlags = MAPI_TOP_LEVEL;
		m_hWnd = hWnd;

		if (!lpszContext.empty())
		{
			m_szContext = lpszContext;
		}
		else
		{
			m_szContext = strings::loadstring(IDS_NOCONTEXT);
		}
	}

	CMAPIProgress::~CMAPIProgress() { TRACE_DESTRUCTOR(CLASS); }

	STDMETHODIMP CMAPIProgress::QueryInterface(REFIID riid, LPVOID* ppvObj)
	{
		*ppvObj = nullptr;
		if (riid == IID_IMAPIProgress || riid == IID_IUnknown)
		{
			*ppvObj = static_cast<LPVOID>(this);
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	STDMETHODIMP_(ULONG) CMAPIProgress::AddRef()
	{
		const auto lCount = InterlockedIncrement(&m_cRef);
		TRACE_ADDREF(CLASS, lCount);
		return lCount;
	}

	STDMETHODIMP_(ULONG) CMAPIProgress::Release()
	{
		const auto lCount = InterlockedDecrement(&m_cRef);
		TRACE_RELEASE(CLASS, lCount);
		if (!lCount) delete this;
		return lCount;
	}

	_Check_return_ STDMETHODIMP CMAPIProgress::Progress(ULONG ulValue, ULONG ulCount, ULONG ulTotal)
	{
		output::DebugPrintEx(
			DBGGeneric,
			CLASS,
			L"Progress",
			L"(%ws) - ulValue = %u, ulCount = %u, ulTotal = %u\n",
			m_szContext.c_str(),
			ulValue,
			ulCount,
			ulTotal);

		OutputState(L"Progress");

		if (m_hWnd)
		{
			const auto iPercent = MulDiv(ulValue - m_ulMin, 100, m_ulMax - m_ulMin);
			dialog::CBaseDialog::UpdateStatus(
				m_hWnd, STATUSINFOTEXT, strings::formatmessage(IDS_PERCENTLOADED, m_szContext.c_str(), iPercent));
		}

		return S_OK;
	}

	STDMETHODIMP CMAPIProgress::GetFlags(ULONG* lpulFlags)
	{
		if (!lpulFlags)
		{
			return MAPI_E_INVALID_PARAMETER;
		}

		OutputState(L"GetFlags");

		*lpulFlags = m_ulFlags;
		return S_OK;
	}

	STDMETHODIMP CMAPIProgress::GetMax(ULONG* lpulMax)
	{
		if (!lpulMax) return MAPI_E_INVALID_PARAMETER;

		OutputState(L"GetMax");

		*lpulMax = m_ulMax;
		return S_OK;
	}

	STDMETHODIMP CMAPIProgress::GetMin(ULONG* lpulMin)
	{
		if (!lpulMin) return MAPI_E_INVALID_PARAMETER;

		OutputState(L"GetMin");

		*lpulMin = m_ulMin;
		return S_OK;
	}

	STDMETHODIMP CMAPIProgress::SetLimits(ULONG* lpulMin, ULONG* lpulMax, ULONG* lpulFlags)
	{
		OutputState(L"SetLimits");

		std::wstring szMin;
		std::wstring szMax;
		std::wstring szFlags;

		if (lpulMin)
		{
			szMin = std::to_wstring(*lpulMin);
		}
		else
		{
			szMin = L"NULL";
		}

		if (lpulMax)
		{
			szMax = std::to_wstring(*lpulMax);
		}
		else
		{
			szMin = L"NULL";
		}

		if (lpulFlags)
		{
			szFlags = strings::format(L"0x%08X", *lpulFlags); // STRING_OK
		}
		else
		{
			szMin = L"NULL";
		}

		output::DebugPrintEx(
			DBGGeneric,
			CLASS,
			L"SetLimits",
			L"(%ws) - Passed Values: lpulMin = %ws, lpulMax = %ws, lpulFlags = %ws\n",
			m_szContext.c_str(),
			szMin.c_str(),
			szMax.c_str(),
			szFlags.c_str());

		if (lpulMin) m_ulMin = *lpulMin;

		if (lpulMax) m_ulMax = *lpulMax;

		if (lpulFlags) m_ulFlags = *lpulFlags;

		OutputState(L"SetLimits");

		return S_OK;
	}

	void CMAPIProgress::OutputState(const std::wstring& lpszFunction) const
	{
		output::DebugPrint(
			DBGGeneric,
			L"%ws::%ws(%ws) - Current Values: Min = %u, Max = %u, Flags = %u\n",
			CLASS.c_str(),
			lpszFunction.c_str(),
			m_szContext.c_str(),
			m_ulMin,
			m_ulMax,
			m_ulFlags);
	}
} // namespace ui