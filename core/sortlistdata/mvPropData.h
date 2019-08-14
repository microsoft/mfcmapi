#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	class mvPropData : public IData
	{
	public:
		static void init(sortListData* data, _In_ const _SPropValue* lpProp, ULONG iProp);
		static void init(sortListData* data, _In_opt_ const _SPropValue* lpProp);

		_PV m_val{};

	private:
		mvPropData(_In_opt_ const _SPropValue* lpProp, ULONG iProp);
		mvPropData(_In_opt_ const _SPropValue* lpProp);

		std::string m_lpszA{};
		std::wstring m_lpszW{};
		std::vector<BYTE> m_lpBin{};
	};
} // namespace sortlistdata