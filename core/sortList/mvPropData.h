#pragma once
#include <MAPIDefS.h>
#include <core/sortList/data.h>

namespace controls
{
	namespace sortlistdata
	{
		class mvPropData : public IData
		{
		public:
			mvPropData(_In_opt_ const _SPropValue* lpProp, ULONG iProp);
			mvPropData(_In_opt_ const _SPropValue* lpProp);
			_PV m_val{};

		private:
			std::string m_lpszA;
			std::wstring m_lpszW;
			std::vector<BYTE> m_lpBin;
		};
	} // namespace sortlistdata
} // namespace controls